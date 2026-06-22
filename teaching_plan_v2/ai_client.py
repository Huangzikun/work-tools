"""
豆包AI客户端 - 基于火山引擎 Ark Responses API (OpenAI SDK)

对外保持 AIClient 接口兼容，内部委托给 common.llm_client.LLMClient。
"""
import logging

from common.llm_client import LLMClient

logger = logging.getLogger(__name__)


class AIClient:
    """豆包AI客户端（接口兼容层，内部委托 LLMClient）"""

    def __init__(
        self,
        api_key: str = None,
        model: str = "doubao-seed-2-0-mini-260428",
        timeout: int = 1800,
        max_retries: int = 2
    ):
        """
        初始化AI客户端

        Args:
            api_key: API密钥，如果为None则从环境变量ARK_API_KEY读取
            model: 模型名称（默认统一为新的 doubao-seed-2-0-mini-260428）
            timeout: 超时时间（秒）
            max_retries: 最大重试次数
        """
        self._llm = LLMClient(
            api_key=api_key,
            model=model,
            timeout=timeout,
            max_retries=max_retries,
        )
        self.model = model
        logger.info(f"AI客户端初始化完成，模型: {model}")

    def generate_response(
        self,
        system_prompt: str,
        user_prompt: str,
        response_format: str = "json"
    ) -> str:
        """
        调用豆包模型生成响应

        Args:
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            response_format: 响应格式 ("json" 或 "text")

        Returns:
            模型返回的文本内容
        """
        json_output = response_format == "json"
        try:
            result = self._llm.generate(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                json_output=json_output,
            )
            logger.info(f"AI调用成功，输出长度: {len(result)} 字符")
            return result
        except Exception as e:
            logger.error(f"AI调用失败: {e}")
            raise

    def generate_response_with_retry(
        self,
        system_prompt: str,
        user_prompt: str,
        max_retries: int = 3,
        response_format: str = "json"
    ) -> str:
        """
        带重试的AI调用

        Args:
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            max_retries: 最大重试次数
            response_format: 响应格式

        Returns:
            模型返回的文本内容
        """
        json_output = response_format == "json"
        return self._llm.generate_with_retry(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            json_output=json_output,
            max_retries=max_retries,
        )
