import json
import os
from datetime import datetime
from typing import Dict, List, Any, Optional
from src.utils.logger import Logger

class HistoryManager:
    """
    历史记录管理器
    管理用户的转换历史记录
    """
    
    def __init__(self, history_file: str = "history.json", max_items: int = 50):
        """
        初始化历史记录管理器
        
        Args:
            history_file: 历史记录文件路径
            max_items: 最大记录数量
        """
        self.history_file = history_file
        self.max_items = max_items
        self.logger = Logger().get_logger()
        self._history: List[Dict[str, Any]] = []
        self._load_history()
    
    def _load_history(self):
        """
        加载历史记录
        """
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    self._history = json.load(f)
                self.logger.info(f"历史记录加载成功，共 {len(self._history)} 条记录")
            except Exception as e:
                self.logger.error(f"加载历史记录失败: {e}")
                self._history = []
        else:
            self.logger.info("历史记录文件不存在，创建新的历史记录")
            self._history = []
            self._save_history()
    
    def _save_history(self):
        """
        保存历史记录
        """
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self._history, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.logger.error(f"保存历史记录失败: {e}")
    
    def add_record(self, input_path: str, output_path: str, mode: str, 
                  output_format: str, success: bool, duration: float = None,
                  file_size: float = None, slide_count: int = None) -> Dict[str, Any]:
        """
        添加历史记录
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            mode: 转换模式
            output_format: 输出格式
            success: 是否成功
            duration: 耗时（秒）
            file_size: 输出文件大小（MB）
            slide_count: 幻灯片数量
            
        Returns:
            Dict[str, Any]: 新添加的记录
        """
        record = {
            "id": len(self._history) + 1,
            "timestamp": datetime.now().isoformat(),
            "input_path": input_path,
            "output_path": output_path,
            "mode": mode,
            "output_format": output_format,
            "success": success,
            "duration": duration,
            "file_size_mb": file_size,
            "slide_count": slide_count
        }
        
        self._history.insert(0, record)
        
        # 限制记录数量
        if len(self._history) > self.max_items:
            self._history = self._history[:self.max_items]
        
        self._save_history()
        self.logger.debug(f"添加历史记录: {input_path}")
        
        return record
    
    def get_all(self) -> List[Dict[str, Any]]:
        """
        获取所有历史记录
        
        Returns:
            List[Dict[str, Any]]: 历史记录列表
        """
        return self._history.copy()
    
    def get_recent(self, count: int = 10) -> List[Dict[str, Any]]:
        """
        获取最近的历史记录
        
        Args:
            count: 记录数量
            
        Returns:
            List[Dict[str, Any]]: 最近的历史记录列表
        """
        return self._history[:count]
    
    def get_by_id(self, record_id: int) -> Optional[Dict[str, Any]]:
        """
        根据ID获取记录
        
        Args:
            record_id: 记录ID
            
        Returns:
            Optional[Dict[str, Any]]: 记录对象，如果不存在则返回None
        """
        for record in self._history:
            if record["id"] == record_id:
                return record.copy()
        return None
    
    def search(self, keyword: str) -> List[Dict[str, Any]]:
        """
        搜索历史记录
        
        Args:
            keyword: 搜索关键词
            
        Returns:
            List[Dict[str, Any]]: 匹配的记录列表
        """
        keyword_lower = keyword.lower()
        results = []
        
        for record in self._history:
            if (keyword_lower in record.get("input_path", "").lower() or
                keyword_lower in record.get("output_path", "").lower() or
                keyword_lower in record.get("mode", "").lower() or
                keyword_lower in record.get("output_format", "").lower()):
                results.append(record)
        
        return results
    
    def filter_by_date(self, start_date: str, end_date: str) -> List[Dict[str, Any]]:
        """
        按日期范围过滤记录
        
        Args:
            start_date: 开始日期（ISO格式）
            end_date: 结束日期（ISO格式）
            
        Returns:
            List[Dict[str, Any]]: 过滤后的记录列表
        """
        results = []
        
        for record in self._history:
            record_date = record.get("timestamp", "")
            if start_date <= record_date <= end_date:
                results.append(record)
        
        return results
    
    def get_statistics(self) -> Dict[str, Any]:
        """
        获取统计信息
        
        Returns:
            Dict[str, Any]: 统计信息字典
        """
        total = len(self._history)
        successful = sum(1 for r in self._history if r.get("success", False))
        failed = total - successful
        
        total_duration = sum(r.get("duration", 0) or 0 for r in self._history)
        avg_duration = total_duration / total if total > 0 else 0
        
        mode_counts = {}
        for record in self._history:
            mode = record.get("mode", "unknown")
            mode_counts[mode] = mode_counts.get(mode, 0) + 1
        
        return {
            "total_records": total,
            "successful": successful,
            "failed": failed,
            "success_rate": successful / total if total > 0 else 0,
            "total_duration": total_duration,
            "average_duration": avg_duration,
            "mode_distribution": mode_counts
        }
    
    def delete_record(self, record_id: int) -> bool:
        """
        删除指定记录
        
        Args:
            record_id: 记录ID
            
        Returns:
            bool: 是否删除成功
        """
        for i, record in enumerate(self._history):
            if record["id"] == record_id:
                del self._history[i]
                self._save_history()
                self.logger.debug(f"删除历史记录: ID={record_id}")
                return True
        return False
    
    def clear_all(self):
        """
        清空所有历史记录
        """
        self._history = []
        self._save_history()
        self.logger.info("已清空所有历史记录")
    
    def clear_old_records(self, days: int = 30):
        """
        清除旧记录
        
        Args:
            days: 保留天数
        """
        cutoff_date = datetime.now().isoformat()[:10]
        cutoff_timestamp = datetime.fromisoformat(cutoff_date).timestamp()
        
        old_count = len(self._history)
        self._history = [
            record for record in self._history
            if datetime.fromisoformat(record["timestamp"]).timestamp() > cutoff_timestamp - (days * 24 * 3600)
        ]
        
        new_count = len(self._history)
        deleted_count = old_count - new_count
        
        if deleted_count > 0:
            self._save_history()
            self.logger.info(f"已清除 {deleted_count} 条旧记录（{days}天前）")
    
    def export_history(self, export_path: str):
        """
        导出历史记录
        
        Args:
            export_path: 导出文件路径
        """
        try:
            with open(export_path, 'w', encoding='utf-8') as f:
                json.dump(self._history, f, indent=4, ensure_ascii=False)
            self.logger.info(f"历史记录已导出到: {export_path}")
        except Exception as e:
            self.logger.error(f"导出历史记录失败: {e}")
            raise
    
    def import_history(self, import_path: str, merge: bool = True):
        """
        导入历史记录
        
        Args:
            import_path: 导入文件路径
            merge: 是否与现有记录合并（False则完全替换）
        """
        try:
            with open(import_path, 'r', encoding='utf-8') as f:
                imported_history = json.load(f)
            
            if merge:
                # 合并记录，避免重复
                existing_ids = {r["id"] for r in self._history}
                for record in imported_history:
                    if record["id"] not in existing_ids:
                        self._history.append(record)
                
                # 重新排序并限制数量
                self._history.sort(key=lambda x: x["timestamp"], reverse=True)
                if len(self._history) > self.max_items:
                    self._history = self._history[:self.max_items]
            else:
                self._history = imported_history
            
            self._save_history()
            self.logger.info(f"历史记录已从 {import_path} 导入")
        except Exception as e:
            self.logger.error(f"导入历史记录失败: {e}")
            raise