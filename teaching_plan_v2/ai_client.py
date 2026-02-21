"""
豆包AI客户端 - 使用Chat Completions API
"""
import os
import time
import logging
from volcenginesdkarkruntime import Ark

logger = logging.getLogger(__name__)


class AIClient:
    """豆包AI客户端"""

    def __init__(
        self,
        api_key: str = None,
        model: str = "doubao-seed-1-6-251015",
        timeout: int = 1800,
        max_retries: int = 2
    ):
        """
        初始化AI客户端

        Args:
            api_key: API密钥，如果为None则从环境变量ARK_API_KEY读取
            model: 模型名称
            timeout: 超时时间（秒）
            max_retries: 最大重试次数
        """
        if api_key is None:
            api_key = os.environ.get("ARK_API_KEY")
            if not api_key:
                raise ValueError("API Key未提供，请设置ARK_API_KEY环境变量或传入api_key参数")

        self.model = model
        self.client = Ark(
            base_url="https://ark.cn-beijing.volces.com/api/v3",
            api_key=api_key,
            timeout=timeout,
            max_retries=max_retries
        )
        logger.info(f"AI客户端初始化完成，模型: {model}")

    def generate_response(
        self,
        system_prompt: str,
        user_prompt: str,
        response_format: str = "json"
    ) -> str:
        """
        使用Chat Completions API调用豆包模型

        Args:
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            response_format: 响应格式 ("json" 或 "text")

        Returns:
            模型返回的文本内容

        Raises:
            Exception: API调用失败时抛出异常
        """
        # 构建messages参数
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]

        logger.debug(f"调用AI模型，消息数量: {len(messages)}")

        try:
            # 调用Chat Completions API
            response = self.client.chat.completions.create(
                model=self.model,
                messages=messages
            )

            # 提取输出内容
            if response.choices and len(response.choices) > 0:
                result = response.choices[0].message.content
                logger.info(f"AI调用成功，输出长度: {len(result)} 字符")
                return result

            logger.warning("AI响应为空")
            return ""

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
        for attempt in range(max_retries):
            try:
                return self.generate_response(system_prompt, user_prompt, response_format)
            except Exception as e:
                if attempt == max_retries - 1:
                    logger.error(f"AI调用失败，已重试{max_retries}次，放弃重试")
                    raise
                wait_time = 2 ** attempt  # 指数退避
                logger.warning(f"AI调用失败，{wait_time}秒后重试 (第{attempt + 1}/{max_retries}次)")
                time.sleep(wait_time)

        # 理论上不会执行到这里
        raise Exception("AI调用失败")
