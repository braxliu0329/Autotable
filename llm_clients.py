import requests
import json
import logging
import time
from ollama import Client

logger = logging.getLogger(__name__)

class BaseLLMClient:
    """大语言模型客户端基类"""
    def chat_completion(self, messages, temperature=0.7):
        raise NotImplementedError("子类必须实现chat_completion方法")

    def chat_with_tools(self, messages, tools=None, temperature=0.7):
        """支持工具调用的聊天方法，返回完整的 message 对象 (content + tool_calls)"""
        raise NotImplementedError("子类必须实现chat_with_tools方法")

class APIClient(BaseLLMClient):
    """OpenAI兼容API客户端"""
    def __init__(self, api_base_url, api_key, model_name, timeout=60):
        self.api_base_url = api_base_url.rstrip("/")
        self.api_key = api_key
        self.model_name = model_name
        self.timeout = timeout
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }

    def _make_request(self, messages, tools=None, temperature=0.7, max_retries=3):
        url = f"{self.api_base_url}/chat/completions"
        payload = {
            "model": self.model_name,
            "messages": messages,
            "temperature": temperature
        }
        if tools:
            payload["tools"] = tools

        last_exception = None
        for attempt in range(max_retries):
            try:
                response = requests.post(url, headers=self.headers, json=payload, timeout=self.timeout)
                
                if 500 <= response.status_code < 600:
                    response.raise_for_status()
                
                response.raise_for_status()
                return response.json()["choices"][0]["message"]
                
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout, requests.exceptions.HTTPError) as e:
                if isinstance(e, requests.exceptions.HTTPError):
                    if not (500 <= e.response.status_code < 600):
                        logger.error(f"API请求客户端错误 (无法重试): {str(e)}")
                        raise
                
                last_exception = e
                if attempt < max_retries - 1:
                    wait_time = 2 ** attempt
                    error_content = ""
                    if hasattr(e, 'response') and e.response is not None:
                        try:
                            error_content = f" | Server Response: {e.response.text[:1000]}"
                        except:
                            pass
                    logger.info(f"API请求异常详细信息: {str(e)}{error_content}")
                    logger.warning(f"API请求失败 ({str(e)})，{wait_time}秒后进行第 {attempt + 1}/{max_retries} 次重试...")
                    time.sleep(wait_time)
                else:
                    logger.error(f"API请求最终失败，已重试 {max_retries} 次: {str(e)}")
            except KeyError as e:
                logger.error(f"API响应格式异常: {str(e)}")
                raise
            except Exception as e:
                logger.error(f"发生未预期错误: {str(e)}")
                raise

        if last_exception:
            raise last_exception

    def chat_completion(self, messages, temperature=0.7, max_retries=3):
        message = self._make_request(messages, temperature=temperature, max_retries=max_retries)
        return message["content"]

    def chat_with_tools(self, messages, tools=None, temperature=0.7, max_retries=3):
        return self._make_request(messages, tools=tools, temperature=temperature, max_retries=max_retries)

class OllamaClient(BaseLLMClient):
    """本地Ollama客户端"""
    def __init__(self, host, model_name, timeout=60):
        self.client = Client(host=host)
        self.model_name = model_name
        self.timeout = timeout

    def chat_completion(self, messages, temperature=0.3):
        try:
            response = self.client.chat(
                model=self.model_name,
                messages=messages,
                options={"temperature": temperature}
            )
            return response['message']['content']
        except Exception as e:
            logger.error(f"Ollama调用失败: {str(e)}")
            raise

    def chat_with_tools(self, messages, tools=None, temperature=0.3):
        try:
            # Check if tools is not None to avoid passing it if library doesn't support it?
            # Assuming recent ollama library supports it.
            kwargs = {"temperature": temperature}
            if tools:
                # Ollama client chat method supports 'tools' param directly, not inside options
                response = self.client.chat(
                    model=self.model_name,
                    messages=messages,
                    tools=tools,
                    options=kwargs
                )
            else:
                response = self.client.chat(
                    model=self.model_name,
                    messages=messages,
                    options=kwargs
                )
            return response['message']
        except Exception as e:
            logger.error(f"Ollama调用失败 (Tools): {str(e)}")
            raise
