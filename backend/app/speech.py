"""Persistent client to the existing Whisper MCP; no ML packages in the gateway."""
import asyncio
import json
import ctypes
import os
from pathlib import Path
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from .llm import ServiceError

class WhisperSpeech:
    def __init__(self, project: Path):
        self.project = project.resolve()
        self.config_path = self.project / 'config.json'
        self.upload_root = self.project / 'audio' / 'madar'
        self.jobs = asyncio.Queue(maxsize=1)
        self.task = None
        self.lock = asyncio.Lock()

    def configuration(self):
        try:
            config = json.loads(self.config_path.read_text('utf-8-sig'))
            if not Path(config['python_executable']).is_file(): raise ValueError()
            if not (Path(config['model_path']) / 'model.bin').is_file(): raise ValueError()
            if not (self.project / 'server.py').is_file(): raise ValueError()
            if not any(self.upload_root.is_relative_to(Path(root).resolve()) for root in config['allowed_roots']): raise ValueError()
            return config
        except (OSError, ValueError, KeyError) as exc:
            raise ServiceError('stt_unavailable', 'تنظیمات، محیط Python یا مدل موجود Whisper آماده نیست؛ هیچ دانلودی انجام نمی‌شود.') from exc

    def health(self):
        try:
            config = self.configuration()
            return {'available': True, 'model': config['model_repository'], 'transport': 'stdio',
                    'formats': ['wav','mp3','ogg'], 'max_file_mb': 25,
                    'status': 'configured_lazy_load', 'busy': self.lock.locked()}
        except ServiceError as exc: return {'available': False, 'message': exc.message}

    def available_ram(self):
        if os.name == 'nt':
            class Memory(ctypes.Structure):
                _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('total','available','page_total','page_available','virtual_total','virtual_available','extended')]
            info=Memory();info.length=ctypes.sizeof(info)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(info)):return info.available
        return None

    async def transcribe(self, path: Path) -> dict:
        self.configuration()
        if self.lock.locked(): raise ServiceError('stt_busy', 'یک تبدیل صوت در حال اجراست؛ کمی صبر کنید.')
        async with self.lock:
            if self.task is None or self.task.done(): self.task = asyncio.create_task(self._serve())
            future = asyncio.get_running_loop().create_future()
            await self.jobs.put((path, future))
            try:
                async with asyncio.timeout(330): return await asyncio.shield(future)
            except (TimeoutError, asyncio.CancelledError):
                # Cancel the session to stop its fixed worker before removing the upload.
                await self.close()
                raise

    async def _serve(self):
        current = None
        try:
            config = self.configuration()
            params = StdioServerParameters(command=config['python_executable'],
                args=['-I','-X','utf8',str(self.project/'server.py'),'--config',str(self.config_path)])
            async with stdio_client(params) as (read, write):
                async with ClientSession(read, write) as session:
                    await session.initialize()
                    listing = await session.list_tools()
                    if not any(t.name=='transcribe_audio' for t in listing.tools): raise ValueError('Missing STT tool')
                    while True:
                        path, current = await self.jobs.get()
                        result = await session.call_tool('transcribe_audio', {'file_path':str(path)})
                        if result.isError:
                            detail='\n'.join(c.text for c in result.content if c.type=='text')[:2000]
                            current.set_exception(ServiceError('stt_failed',detail or 'تبدیل صوت ناموفق بود.'))
                        else:
                            data=result.structuredContent
                            if not isinstance(data,dict) or not isinstance(data.get('text'),str) or not data['text'].strip():
                                current.set_exception(ServiceError('stt_empty','Whisper متن واقعی برنگرداند.'))
                            else: current.set_result(data)
                        current=None
        except asyncio.CancelledError:
            if current is not None and not current.done(): current.cancel()
            raise
        except Exception:
            error=ServiceError('stt_connection_failed','ارتباط با Whisper موجود کامل نشد؛ هیچ مدلی دانلود نشد.')
            if current is not None and not current.done(): current.set_exception(error)
            while not self.jobs.empty():
                _, pending=self.jobs.get_nowait()
                if not pending.done(): pending.set_exception(error)

    async def close(self):
        if self.task is not None:
            self.task.cancel()
            try: await self.task
            except asyncio.CancelledError: pass
            self.task=None
        while not self.jobs.empty():
            _, pending=self.jobs.get_nowait()
            if not pending.done():pending.cancel()

def audio_signature(header: bytes, extension: str) -> bool:
    if extension=='.wav': return header[:4]==b'RIFF' and header[8:12]==b'WAVE'
    if extension=='.ogg': return header.startswith(b'OggS')
    if extension=='.mp3': return header.startswith(b'ID3') or (len(header)>1 and header[0]==255 and header[1]&224==224)
    return False
