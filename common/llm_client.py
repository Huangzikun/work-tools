"""
统一 LLM 客户端 - 基于火山引擎 Ark Responses API (OpenAI SDK)

快速开始:
    from common.llm_client import LLMClient

    client = LLMClient()  # 自动从 ARK_API_KEY 环境变量读取

    # 文本对话
    text = client.generate(system_prompt="你是助理", user_prompt="你好")

    # JSON 输出
    json_text = client.generate(
        system_prompt="你是评分老师",
        user_prompt="...",
        json_output=True,
    )

    # 多模态（图片 / 文件）
    multi_text = client.generate(
        system_prompt="...",
        user_content=[
            {"type": "input_file", "file_id": file_id},
            {"type": "input_text", "text": "..."},
        ],
        json_schema={...},
        thinking_disabled=True,
    )
"""

import os
import time
import logging
from typing import Optional, List, Dict, Any

from openai import OpenAI

logger = logging.getLogger(__name__)

DEFAULT_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
DEFAULT_MODEL = "doubao-seed-2-0-mini-260428"


class LLMClient:
    """火山引擎 Ark 统一大模型客户端，基于 OpenAI SDK 的 Responses API。"""

    def __init__(
        self,
        api_key: Optional[str] = None,
        base_url: str = DEFAULT_BASE_URL,
        model: str = DEFAULT_MODEL,
        timeout: int = 1800,
        max_retries: int = 2,
    ):
        """
        Args:
            api_key: API 密钥，None 时从 ARK_API_KEY 环境变量读取
            base_url: 火山引擎 Ark 服务地址
            model: 默认模型名称
            timeout: 单次请求超时秒数
            max_retries: SDK 层面的重试次数
        """
        if api_key is None:
            api_key = os.environ.get("ARK_API_KEY")
            if not api_key:
                raise ValueError(
                    "API Key 未提供，请设置 ARK_API_KEY 环境变量或传入 api_key 参数"
                )

        self.api_key = api_key
        self.base_url = base_url
        self.model = model
        self.timeout = timeout
        self.max_retries = max_retries

        self.client = OpenAI(
            base_url=base_url,
            api_key=api_key,
            timeout=timeout,
            max_retries=max_retries,
        )
        logger.info(f"LLM 客户端初始化完成，模型: {model}")

    @staticmethod
    def _build_content(content) -> List[Dict[str, Any]]:
        """
        将内容转换为 Responses API 多模态格式：
        - 字符串 → [{"type": "input_text", "text": content}]
        - list → 原样返回（支持 input_text / input_image / input_file）
        """
        if isinstance(content, str):
            return [{"type": "input_text", "text": content}]
        if isinstance(content, list):
            return content
        raise TypeError(f"不支持的 content 类型: {type(content)}")

    def _build_input(
        self,
        system_prompt: Optional[str],
        user_prompt: Optional[str],
        user_content: Optional[List[Dict[str, Any]]],
    ) -> List[Dict[str, Any]]:
        messages: List[Dict[str, Any]] = []
        if system_prompt:
            messages.append(
                {
                    "role": "system",
                    "content": self._build_content(system_prompt),
                }
            )
        if user_content is not None:
            messages.append({"role": "user", "content": user_content})
        elif user_prompt is not None:
            messages.append(
                {
                    "role": "user",
                    "content": self._build_content(user_prompt),
                }
            )
        return messages

    def generate(
        self,
        system_prompt: Optional[str] = None,
        user_prompt: Optional[str] = None,
        user_content: Optional[List[Dict[str, Any]]] = None,
        json_output: bool = False,
        json_schema: Optional[Dict[str, Any]] = None,
        max_tokens: Optional[int] = None,
        thinking_disabled: bool = False,
    ) -> str:
        """
        调用 responses.create，返回文本输出。

        Args:
            system_prompt: 系统提示词字符串
            user_prompt: 用户提示词字符串（与 user_content 二选一）
            user_content: 用户消息 content 列表（多模态场景，与 user_prompt 二选一）
            json_output: True 时强制 JSON 输出（json_object 格式）
            json_schema: 传入时使用 json_schema 严格模式（与 json_output 互斥）
            max_tokens: 最大输出 token 数
            thinking_disabled: True 时关闭深度思考

        Returns:
            模型返回的文本内容
        """
        input_messages = self._build_input(system_prompt, user_prompt, user_content)

        kwargs: Dict[str, Any] = {
            "model": self.model,
            "input": input_messages,
        }

        if json_schema is not None:
            kwargs["text"] = {
                "format": {
                    "type": "json_schema",
                    "name": "output",
                    "strict": True,
                    "schema": json_schema,
                }
            }
        elif json_output:
            kwargs["text"] = {"format": {"type": "json_object"}}

        if max_tokens is not None:
            kwargs["max_output_tokens"] = max_tokens

        # thinking 是火山引擎 Ark 扩展字段，OpenAI SDK 不会透传，
        # 必须通过 extra_body 注入到 HTTP 请求体中
        if thinking_disabled:
            kwargs["extra_body"] = {"thinking": {"type": "disabled"}}

        logger.debug(
            f"调用 Responses API，消息数量: {len(input_messages)}，"
            f"json_output={json_output}, json_schema={json_schema is not None}, "
            f"max_tokens={max_tokens}, thinking_disabled={thinking_disabled}"
        )

        try:
            response = self.client.responses.create(**kwargs)
            output_text = response.output_text or ""
            logger.info(f"LLM 调用成功，输出长度: {len(output_text)} 字符")
            return output_text
        except Exception as e:
            logger.error(f"LLM 调用失败: {e}")
            raise

    def generate_with_retry(
        self,
        system_prompt: Optional[str] = None,
        user_prompt: Optional[str] = None,
        user_content: Optional[List[Dict[str, Any]]] = None,
        json_output: bool = False,
        json_schema: Optional[Dict[str, Any]] = None,
        max_tokens: Optional[int] = None,
        thinking_disabled: bool = False,
        max_retries: int = 3,
    ) -> str:
        """带指数退避重试的调用。"""
        for attempt in range(max_retries):
            try:
                return self.generate(
                    system_prompt=system_prompt,
                    user_prompt=user_prompt,
                    user_content=user_content,
                    json_output=json_output,
                    json_schema=json_schema,
                    max_tokens=max_tokens,
                    thinking_disabled=thinking_disabled,
                )
            except Exception as e:
                if attempt == max_retries - 1:
                    logger.error(f"LLM 调用失败，已重试 {max_retries} 次，放弃重试")
                    raise
                wait_time = 2 ** attempt
                logger.warning(
                    f"LLM 调用失败，{wait_time} 秒后重试 (第 {attempt + 1}/{max_retries} 次)，原因: {e}"
                )
                time.sleep(wait_time)
        raise Exception("LLM 调用失败")


_default_client: Optional[LLMClient] = None


def get_default_client() -> LLMClient:
    """获取进程级单例 LLMClient，避免每个脚本重复初始化。"""
    global _default_client
    if _default_client is None:
        _default_client = LLMClient()
    return _default_client


__all__ = ["LLMClient", "get_default_client", "DEFAULT_BASE_URL", "DEFAULT_MODEL"]
