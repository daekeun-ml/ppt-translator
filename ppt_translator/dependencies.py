"""
Dependency management with graceful fallbacks
"""
import os
import logging
import subprocess
import sys
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)


class DependencyManager:
    """Manages optional dependencies with graceful fallbacks"""
    
    _instance = None
    _initialized = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not self._initialized:
            self.dependencies: Dict[str, Any] = {}
            self._check_dependencies()
            DependencyManager._initialized = True
    
    def _check_dependencies(self):
        """Check and import dependencies"""
        deps = {
            'pptx': ('pptx', 'Presentation'),
            'boto3': ('boto3', None),
            'dotenv': ('dotenv', 'load_dotenv'),
            'mcp': ('mcp', 'MCPServer')
        }
        
        for key, (module_name, attr) in deps.items():
            try:
                module = __import__(module_name)
                if attr:
                    self.dependencies[key] = getattr(module, attr) if hasattr(module, attr) else module
                else:
                    self.dependencies[key] = module
                logger.debug(f"✅ {module_name} loaded successfully")
            except ImportError as e:
                logger.warning(f"⚠️ {module_name} not available: {e}")
                self.dependencies[key] = None
    
    def get(self, key: str) -> Optional[Any]:
        """Get dependency or None if not available"""
        return self.dependencies.get(key)
    
    def is_available(self, key: str) -> bool:
        """Check if dependency is available"""
        return self.dependencies.get(key) is not None
    
    def require(self, key: str) -> Any:
        """Require dependency or raise error"""
        dep = self.dependencies.get(key)
        if dep is None:
            raise ImportError(f"Required dependency '{key}' is not available")
        return dep


def install_dependencies():
    """Install required dependencies using pip"""
    packages = ["mcp", "python-pptx", "boto3", "python-dotenv"]
    
    try:
        logger.info("📦 Installing dependencies with pip...")
        cmd = [sys.executable, "-m", "pip", "install"] + packages
        subprocess.check_call(cmd)
        logger.info("✅ Dependencies installed successfully!")
        
    except Exception as e:
        logger.error(f"❌ Failed to install dependencies: {str(e)}")
        sys.exit(1)
