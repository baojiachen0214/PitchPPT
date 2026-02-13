import json
import os
from typing import Dict, Any, Optional
from src.utils.logger import Logger

class ConfigManager:
    """
    配置文件管理器
    支持JSON格式的配置文件读写
    """
    
    def __init__(self, config_file: str = "config.json"):
        """
        初始化配置管理器
        
        Args:
            config_file: 配置文件路径
        """
        self.config_file = config_file
        self.logger = Logger().get_logger()
        self._config = {}
        self._load_config()
    
    def _load_config(self):
        """
        加载配置文件
        """
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self._config = json.load(f)
                self.logger.info(f"配置文件加载成功: {self.config_file}")
            except Exception as e:
                self.logger.error(f"加载配置文件失败: {e}")
                self._config = self._get_default_config()
        else:
            self.logger.info("配置文件不存在，使用默认配置")
            self._config = self._get_default_config()
            self._save_config()
    
    def _save_config(self):
        """
        保存配置文件
        """
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=4, ensure_ascii=False)
            self.logger.info(f"配置文件保存成功: {self.config_file}")
        except Exception as e:
            self.logger.error(f"保存配置文件失败: {e}")
    
    def _get_default_config(self) -> Dict[str, Any]:
        """
        获取默认配置
        
        Returns:
            Dict[str, Any]: 默认配置字典
        """
        return {
            "app": {
                "name": "PitchPPT",
                "version": "1.0.0",
                "author": "Jiachen Bao",
                "language": "zh-CN"
            },
            "conversion": {
                "default_output_format": "pptx",
                "default_image_quality": 95,
                "default_resolution_scale": 1.0,
                "preserve_aspect_ratio": True,
                "include_hidden_slides": False
            },
            "ui": {
                "theme": "default",
                "window_width": 1200,
                "window_height": 800,
                "remember_window_size": True,
                "show_file_info": True
            },
            "logging": {
                "level": "INFO",
                "max_file_size_mb": 10,
                "backup_count": 5,
                "cleanup_days": 30
            },
            "paths": {
                "default_output_dir": "",
                "last_input_dir": "",
                "last_output_dir": "",
                "template_dir": "templates",
                "history_file": "history.json"
            },
            "advanced": {
                "enable_performance_logging": True,
                "enable_auto_cleanup": True,
                "max_history_items": 50
            }
        }
    
    def get(self, key: str, default: Any = None) -> Any:
        """
        获取配置值
        
        Args:
            key: 配置键（支持点号分隔的嵌套键，如 "conversion.default_mode"）
            default: 默认值
            
        Returns:
            Any: 配置值
        """
        keys = key.split('.')
        value = self._config
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def set(self, key: str, value: Any, save: bool = True):
        """
        设置配置值
        
        Args:
            key: 配置键（支持点号分隔的嵌套键）
            value: 配置值
            save: 是否立即保存到文件
        """
        keys = key.split('.')
        config = self._config
        
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        
        config[keys[-1]] = value
        
        if save:
            self._save_config()
        
        self.logger.debug(f"配置更新: {key} = {value}")
    
    def get_all(self) -> Dict[str, Any]:
        """
        获取所有配置
        
        Returns:
            Dict[str, Any]: 完整的配置字典
        """
        return self._config.copy()
    
    def update(self, config: Dict[str, Any], save: bool = True):
        """
        批量更新配置
        
        Args:
            config: 配置字典
            save: 是否立即保存到文件
        """
        self._deep_update(self._config, config)
        
        if save:
            self._save_config()
        
        self.logger.info("配置批量更新完成")
    
    def update_config(self, config: Dict[str, Any], save: bool = True):
        """
        更新配置（别名方法，与 update 相同）
        
        Args:
            config: 配置字典
            save: 是否立即保存到文件
        """
        return self.update(config, save)
    
    def _deep_update(self, base: Dict, update: Dict):
        """
        深度更新字典
        
        Args:
            base: 基础字典
            update: 更新字典
        """
        for key, value in update.items():
            if isinstance(value, dict) and key in base and isinstance(base[key], dict):
                self._deep_update(base[key], value)
            else:
                base[key] = value
    
    def reset_to_default(self, save: bool = True):
        """
        重置为默认配置
        
        Args:
            save: 是否立即保存到文件
        """
        self._config = self._get_default_config()
        
        if save:
            self._save_config()
        
        self.logger.info("配置已重置为默认值")
    
    def export_config(self, export_path: str):
        """
        导出配置到指定路径
        
        Args:
            export_path: 导出文件路径
        """
        try:
            with open(export_path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=4, ensure_ascii=False)
            self.logger.info(f"配置已导出到: {export_path}")
        except Exception as e:
            self.logger.error(f"导出配置失败: {e}")
            raise
    
    def import_config(self, import_path: str, merge: bool = True):
        """
        从指定路径导入配置
        
        Args:
            import_path: 导入文件路径
            merge: 是否与现有配置合并（False则完全替换）
        """
        try:
            with open(import_path, 'r', encoding='utf-8') as f:
                imported_config = json.load(f)
            
            if merge:
                self._deep_update(self._config, imported_config)
            else:
                self._config = imported_config
            
            self._save_config()
            self.logger.info(f"配置已从 {import_path} 导入")
        except Exception as e:
            self.logger.error(f"导入配置失败: {e}")
            raise
    
    def validate_config(self) -> bool:
        """
        验证配置的有效性
        
        Returns:
            bool: 配置是否有效
        """
        try:
            # 检查必需的顶级键
            required_keys = ["app", "conversion", "ui", "logging", "paths", "advanced"]
            for key in required_keys:
                if key not in self._config:
                    self.logger.error(f"缺少必需的配置键: {key}")
                    return False
            
            # 检查图片质量范围
            quality = self.get("conversion.default_image_quality", 95)
            if not 1 <= quality <= 100:
                self.logger.error(f"图片质量超出范围: {quality}")
                return False
            
            # 检查分辨率缩放
            scale = self.get("conversion.default_resolution_scale", 1.0)
            if not 0.1 <= scale <= 5.0:
                self.logger.error(f"分辨率缩放超出范围: {scale}")
                return False
            
            # 检查日志级别
            level = self.get("logging.level", "INFO")
            if level not in ["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]:
                self.logger.error(f"无效的日志级别: {level}")
                return False
            
            self.logger.info("配置验证通过")
            return True
            
        except Exception as e:
            self.logger.error(f"配置验证失败: {e}")
            return False