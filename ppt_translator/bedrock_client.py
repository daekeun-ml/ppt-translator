"""
AWS Bedrock client wrapper with connection management
"""
import os
import logging
from typing import Optional, Any
from .dependencies import DependencyManager

logger = logging.getLogger(__name__)


class BedrockClient:
    """AWS Bedrock client wrapper with connection management"""
    
    def __init__(self, region: str = None):
        self._client = None
        self._initialized = False
        self.region = region or os.getenv('AWS_REGION', 'us-east-1')
        self.deps = DependencyManager()
    
    @property
    def client(self) -> Optional[Any]:
        """Lazy initialization of Bedrock client"""
        if not self._initialized:
            self._initialize()
        return self._client
    
    def _initialize(self) -> bool:
        """Initialize the AWS Bedrock client"""
        try:
            boto3 = self.deps.require('boto3')
            logger.info(f"Initializing Bedrock client with region: {self.region}")
            
            # Try default credential chain first
            try:
                self._client = boto3.client('bedrock-runtime', region_name=self.region)
                logger.info("✅ Bedrock client initialized with default credentials")
                self._initialized = True
                return True
            except Exception as e:
                logger.warning(f"Default credentials failed: {str(e)}")
            
            # Fallback to explicit credentials
            access_key = os.getenv('AWS_ACCESS_KEY_ID')
            secret_key = os.getenv('AWS_SECRET_ACCESS_KEY')
            
            if access_key and secret_key and not access_key.startswith('${'):
                self._client = boto3.client(
                    'bedrock-runtime',
                    aws_access_key_id=access_key,
                    aws_secret_access_key=secret_key,
                    region_name=self.region
                )
                logger.info("✅ Bedrock client initialized with explicit credentials")
                self._initialized = True
                return True
            else:
                logger.error("❌ AWS credentials not properly configured")
                return False
                
        except Exception as e:
            logger.error(f"❌ Failed to initialize AWS Bedrock client: {str(e)}")
            return False
    
    def is_ready(self) -> bool:
        """Check if client is ready"""
        return self.client is not None
    
    def converse(self, **kwargs) -> Any:
        """Wrapper for converse API call"""
        if not self.is_ready():
            raise Exception("AWS Bedrock client not initialized")
        return self.client.converse(**kwargs)
