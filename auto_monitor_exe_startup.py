#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动监控EXE启动过程 - 自动检测窗口出现并停止监控
"""

import subprocess
import time
import os
import sys
import json
from pathlib import Path
from datetime import datetime
from collections import defaultdict

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False
    print("[错误] 需要安装 psutil: pip install psutil")
    sys.exit(1)

try:
    import win32gui
    import win32process
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("[警告] win32gui 未安装，将无法自动检测窗口出现")
    print("   安装: pip install pywin32")

class AutoProcessMonitor:
    """自动进程监控器 - 自动检测窗口出现"""
    
    def __init__(self, exe_path):
        self.exe_path = Path(exe_path)
        self.start_time = None
        self.main_process = None
        self.processes = {}
        self.process_timeline = []
        self.window_appeared_time = None
        self.monitoring = False
        
    def find_window_by_process(self, pid):
        """根据进程ID查找窗口"""
        if not WIN32_AVAILABLE:
            return None

        def callback(hwnd, windows):
            try:
                # Check both visible and invisible windows
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid:
                    windows.append(hwnd)
            except:
                pass
            return True

        windows = []
        try:
            win32gui.EnumWindows(callback, windows)
            if windows:
                print(f"[调试] 找到 {len(windows)} 个属于PID {pid}的窗口")
                # Return the first window found
                return windows[0]
            else:
                print(f"[调试] 未找到属于PID {pid}的窗口")
                return None
        except Exception as e:
            print(f"[调试] 窗口枚举失败: {e}")
            return None
    
    def wait_for_window(self, pid, timeout=60):
        """等待窗口出现"""
        if not WIN32_AVAILABLE:
            # 如果没有win32，使用固定延迟
            time.sleep(5)
            return time.perf_counter() - self.start_time
        
        start_wait = time.perf_counter()
        check_interval = 0.1
        
        while time.perf_counter() - start_wait < timeout:
            hwnd = self.find_window_by_process(pid)
            if hwnd:
                elapsed = time.perf_counter() - self.start_time
                self.window_appeared_time = elapsed
                self.log_event("WINDOW_APPEARED", {
                    'pid': pid,
                    'hwnd': hwnd,
                    'elapsed_seconds': elapsed
                })
                return elapsed
            time.sleep(check_interval)
        
        return None
    
    def start_monitoring(self, max_duration=60):
        """开始监控"""
        print(f"开始监控: {self.exe_path.name}")
        print("=" * 60)
        
        self.start_time = time.perf_counter()
        self.monitoring = True
        
        # 启动主进程
        try:
            self.main_process = subprocess.Popen(
                [str(self.exe_path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
            )
            
            main_pid = self.main_process.pid
            elapsed = time.perf_counter() - self.start_time
            
            self.log_event("MAIN_PROCESS_STARTED", {
                'pid': main_pid,
                'exe_path': str(self.exe_path),
                'elapsed_seconds': elapsed
            })
            
            print(f"[{elapsed:.3f}s] 主进程启动 PID: {main_pid}")
            
            # 监控进程树
            seen_pids = set()
            self.scan_process_tree(main_pid, seen_pids)
            
            # 等待窗口出现
            print("\n等待窗口出现...")
            window_time = self.wait_for_window(main_pid, timeout=max_duration)
            
            if window_time:
                print(f"\n[{window_time:.3f}s] 窗口已出现！")
                # 窗口出现后继续监控一段时间
                time.sleep(2)
            else:
                print("\n[超时] 未检测到窗口，继续监控...")
            
            # 继续监控一段时间以捕获所有进程
            print("\n继续监控进程...")
            end_time = time.perf_counter() + 5  # 再监控5秒
            
            while time.perf_counter() < end_time:
                self.scan_process_tree(main_pid, seen_pids)
                time.sleep(0.2)
            
            print("\n监控完成")
            
        except Exception as e:
            print(f"[错误] 无法启动进程: {e}")
            return False
        
        return True
    
    def log_event(self, event_type, data):
        """记录事件"""
        elapsed = time.perf_counter() - self.start_time
        event = {
            'time': elapsed,
            'event_type': event_type,
            'data': data
        }
        self.process_timeline.append(event)
    
    def scan_process_tree(self, root_pid, seen_pids):
        """扫描进程树"""
        try:
            root_process = psutil.Process(root_pid)
            children = root_process.children(recursive=True)
            
            # 包括主进程
            all_processes = [root_process] + children
            
            for proc in all_processes:
                pid = proc.pid
                
                if pid not in seen_pids:
                    seen_pids.add(pid)
                    self.record_process_info(proc, "主进程" if pid == root_pid else f"子进程-{pid}")
        
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    
    def record_process_info(self, process, label):
        """记录进程详细信息"""
        try:
            pid = process.pid
            elapsed = time.perf_counter() - self.start_time
            
            try:
                proc_info = process.as_dict([
                    'pid', 'name', 'exe', 'cmdline', 'create_time',
                    'cpu_percent', 'memory_info', 'num_threads', 'status'
                ])
            except:
                proc_info = {
                    'pid': pid,
                    'name': process.name(),
                    'exe': None
                }
            
            # 计算进程启动时间
            if 'create_time' in proc_info and proc_info['create_time']:
                try:
                    proc_create_time = proc_info['create_time']
                    # psutil returns create_time as seconds since epoch, convert to process age
                    proc_elapsed = time.time() - proc_create_time
                except:
                    proc_elapsed = elapsed
            else:
                proc_elapsed = elapsed
            
            # 获取资源使用
            try:
                cpu_percent = process.cpu_percent(interval=0.05)
                mem_info = process.memory_info()
            except:
                cpu_percent = 0
                mem_info = type('obj', (object,), {'rss': 0, 'vms': 0})()
            
            process_data = {
                'pid': pid,
                'label': label,
                'name': proc_info.get('name', 'unknown'),
                'exe': proc_info.get('exe', ''),
                'cmdline': proc_info.get('cmdline', []),
                'discovered_at': elapsed,
                'process_age': proc_elapsed,
                'cpu_percent': cpu_percent,
                'memory_mb': mem_info.rss / 1024 / 1024,
                'num_threads': proc_info.get('num_threads', 0),
                'status': proc_info.get('status', 'unknown')
            }
            
            self.processes[pid] = process_data
            self.log_event("PROCESS_DISCOVERED", process_data)
            
            print(f"[{elapsed:.3f}s] 发现进程: {proc_info.get('name', 'unknown')} (PID: {pid})")
            
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    
    def generate_report(self):
        """生成详细报告"""
        total_time = time.perf_counter() - self.start_time
        
        sorted_processes = sorted(
            self.processes.values(),
            key=lambda x: x['discovered_at']
        )
        
        report = {
            'exe_path': str(self.exe_path),
            'analysis_time': datetime.now().isoformat(),
            'total_monitoring_time': total_time,
            'window_appeared_time': self.window_appeared_time,
            'time_to_window': self.window_appeared_time if self.window_appeared_time else None,
            'main_process_pid': self.main_process.pid if self.main_process else None,
            'total_processes': len(self.processes),
            'process_timeline': self.process_timeline,
            'processes_by_startup_order': sorted_processes,
            'process_summary': self.generate_summary(sorted_processes)
        }
        
        return report
    
    def generate_summary(self, sorted_processes):
        """生成进程摘要"""
        summary = {
            'first_process': sorted_processes[0] if sorted_processes else None,
            'process_count_by_name': defaultdict(int),
            'total_memory_mb': 0,
            'peak_cpu_percent': 0,
            'process_startup_intervals': [],
            'startup_phases': []
        }

        prev_time = 0
        for proc in sorted_processes:
            summary['process_count_by_name'][proc['name']] += 1
            summary['total_memory_mb'] += proc['memory_mb']
            summary['peak_cpu_percent'] = max(summary['peak_cpu_percent'], proc['cpu_percent'])

            if prev_time > 0:
                interval = proc['discovered_at'] - prev_time
                summary['process_startup_intervals'].append({
                    'from': prev_time,
                    'to': proc['discovered_at'],
                    'interval': interval,
                    'process': proc['name']
                })
            prev_time = proc['discovered_at']

        # 分析启动阶段
        self._analyze_startup_phases(summary, sorted_processes)

        return summary

    def _analyze_startup_phases(self, summary, sorted_processes):
        """分析启动阶段耗时"""
        if not self.process_timeline:
            return

        phases = []

        # 阶段1: 进程启动
        main_process_event = next((e for e in self.process_timeline if e['event_type'] == 'MAIN_PROCESS_STARTED'), None)
        if main_process_event:
            phases.append({
                'phase': '进程启动',
                'start_time': 0,
                'end_time': main_process_event['time'],
                'duration': main_process_event['time'],
                'description': '从监控开始到主进程启动完成'
            })

        # 阶段2: 主进程初始化
        first_process_event = next((e for e in self.process_timeline if e['event_type'] == 'PROCESS_DISCOVERED'), None)
        if first_process_event and main_process_event:
            phases.append({
                'phase': '主进程初始化',
                'start_time': main_process_event['time'],
                'end_time': first_process_event['time'],
                'duration': first_process_event['time'] - main_process_event['time'],
                'description': '主进程启动到进程被发现'
            })

        # 阶段3: 资源加载
        if self.window_appeared_time and first_process_event:
            phases.append({
                'phase': '资源加载',
                'start_time': first_process_event['time'],
                'end_time': self.window_appeared_time,
                'duration': self.window_appeared_time - first_process_event['time'],
                'description': '进程初始化到窗口出现'
            })

        # 阶段4: 界面渲染
        if self.window_appeared_time:
            phases.append({
                'phase': '界面渲染',
                'start_time': self.window_appeared_time,
                'end_time': self.window_appeared_time + 2,  # 假设渲染需要2秒
                'duration': 2,
                'description': '窗口出现到界面完全渲染'
            })

        # 阶段5: 子进程加载
        if len(self.processes) > 1 and self.window_appeared_time:
            last_process = sorted_processes[-1]
            phases.append({
                'phase': '子进程加载',
                'start_time': self.window_appeared_time,
                'end_time': last_process['discovered_at'],
                'duration': last_process['discovered_at'] - self.window_appeared_time,
                'description': '窗口出现到最后子进程加载完成'
            })

        summary['startup_phases'] = phases
    
    def print_report(self, report):
        """打印报告"""
        print("\n" + "=" * 60)
        print("EXE启动过程详细分析报告")
        print("=" * 60)
        
        print(f"\n总监控时间: {report['total_monitoring_time']:.3f} 秒")
        if report['window_appeared_time']:
            print(f"窗口出现时间: {report['window_appeared_time']:.3f} 秒")
            print(f"从启动到界面出现: {report['window_appeared_time']:.3f} 秒")
        print(f"主进程 PID: {report['main_process_pid']}")
        print(f"发现的进程总数: {report['total_processes']}")
        
        print("\n" + "-" * 60)
        print("进程启动顺序 (按发现时间):")
        print("-" * 60)
        
        for i, proc in enumerate(report['processes_by_startup_order'], 1):
            print(f"\n{i}. [{proc['discovered_at']:.3f}s] {proc['name']} (PID: {proc['pid']})")
            print(f"   标签: {proc['label']}")
            if proc['exe']:
                exe_name = Path(proc['exe']).name if proc['exe'] else 'N/A'
                print(f"   可执行文件: {exe_name}")
            print(f"   内存: {proc['memory_mb']:.2f} MB")
            print(f"   CPU: {proc['cpu_percent']:.1f}%")
            print(f"   线程数: {proc['num_threads']}")
            print(f"   状态: {proc['status']}")
        
        print("\n" + "-" * 60)
        print("进程启动时间间隔:")
        print("-" * 60)
        
        intervals = report['process_summary']['process_startup_intervals']
        for interval in intervals:
            print(f"[{interval['from']:.3f}s -> {interval['to']:.3f}s] "
                  f"间隔: {interval['interval']:.3f}s - {interval['process']}")
        
        print("\n" + "-" * 60)
        print("进程统计:")
        print("-" * 60)
        
        for name, count in sorted(report['process_summary']['process_count_by_name'].items()):
            print(f"  {name}: {count} 个进程")
        
        print(f"\n总内存使用: {report['process_summary']['total_memory_mb']:.2f} MB")
        print(f"峰值CPU使用: {report['process_summary']['peak_cpu_percent']:.1f}%")

        # 显示启动阶段分析
        if report['process_summary']['startup_phases']:
            print("\n" + "-" * 60)
            print("启动阶段详细耗时分析:")
            print("-" * 60)

            total_phase_time = 0
            for i, phase in enumerate(report['process_summary']['startup_phases'], 1):
                print(f"\n{i}. [{phase['start_time']:.3f}s -> {phase['end_time']:.3f}s] {phase['phase']}")
                print(f"   持续时间: {phase['duration']:.3f} 秒")
                print(f"   描述: {phase['description']}")
                total_phase_time += phase['duration']

            print(f"\n阶段总耗时: {total_phase_time:.3f} 秒")
            if report['window_appeared_time']:
                efficiency = (total_phase_time / report['window_appeared_time']) * 100 if report['window_appeared_time'] > 0 else 0
                print(f"阶段覆盖率: {efficiency:.1f}%")


def auto_monitor_exe_startup(exe_path, output_file=None):
    """自动监控EXE启动过程"""
    exe_path = Path(exe_path)
    
    if not exe_path.exists():
        print(f"[错误] EXE文件不存在: {exe_path}")
        return None
    
    monitor = AutoProcessMonitor(exe_path)
    
    if not monitor.start_monitoring():
        return None
    
    # 生成报告
    report = monitor.generate_report()
    monitor.print_report(report)
    
    # 保存报告
    if output_file is None:
        output_file = f"exe_startup_auto_monitor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    
    print(f"\n详细报告已保存到: {output_file}")
    
    return report


if __name__ == "__main__":
    # 设置输出编码
    if sys.platform == 'win32':
        try:
            import io
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        except:
            pass
    
    exe_path = Path("dist/correction_optimized/correction_optimized.exe")
    
    if len(sys.argv) > 1:
        exe_path = Path(sys.argv[1])
    
    print("EXE Startup Auto Monitor Tool")
    print("=" * 60)
    print(f"Target EXE: {exe_path}")
    print("\nNote: Will auto-start and monitor, continue 5s after window appears")
    print("=" * 60)
    
    auto_monitor_exe_startup(exe_path)
