"""Bounded, process-local development store. No connection to the business database."""
import secrets
from dataclasses import dataclass, field
from .llm import ServiceError

@dataclass
class Conversation:
    owner: str
    messages: list[dict] = field(default_factory=list)
    evidence: list[dict] = field(default_factory=list)

class MemoryStore:
    def __init__(self): self.conversations: dict[str, Conversation] = {}

    def create(self, owner: str) -> str:
        if len(self.conversations) >= 100:
            raise ServiceError('storage_limit', 'ظرفیت حافظه توسعه پر است؛ راه‌اندازی مجدد لازم است.')
        key = secrets.token_urlsafe(24)
        self.conversations[key] = Conversation(owner)
        return key

    def get(self, key: str, owner: str) -> Conversation:
        item = self.conversations.get(key)
        if item is None or not secrets.compare_digest(item.owner, owner):
            raise ServiceError('conversation_not_found', 'گفت‌وگو پیدا نشد یا متعلق به این نشست نیست.')
        return item
