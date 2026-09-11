"""Optional persistent storage backend for signature images via Vercel Blob.

Local disk (SIGNATURES_DIR/TEMP_SIGNATURES_DIR in renderer.py) works fine on a normal
server with a persistent filesystem, but breaks on Vercel: every serverless invocation
gets a fresh, isolated /tmp, so a signature uploaded in one request is gone by the next.
Vercel automatically sets BLOB_READ_WRITE_TOKEN when Blob storage is enabled on the
project -- its presence is what we use to decide which backend to use, so local dev
(no token set) keeps working exactly as before, untouched.

Two flavors of each operation are exposed:
  - blob_put/blob_get/blob_delete: plain sync functions (asyncio.run() under the hood) --
    safe to call from renderer.py's sync code, which only ever runs from a *sync* FastAPI
    route (FastAPI runs those in a worker thread with no event loop already running).
  - blob_put_async/blob_get_async/blob_delete_async: real coroutines, for the *async*
    upload endpoints in main.py to `await` directly. Those endpoints already run on the
    request's own event loop, so calling asyncio.run() (which starts a brand new loop)
    from inside them raises "asyncio.run() cannot be called from a running event loop" --
    that was a real bug hit in production the sync versions were used there.
"""
import asyncio
import os

BLOB_ENABLED = bool(os.environ.get("BLOB_READ_WRITE_TOKEN"))


async def blob_put_async(pathname: str, data: bytes) -> None:
    from vercel.blob import AsyncBlobClient

    client = AsyncBlobClient()
    await client.put(pathname, data, access="public", add_random_suffix=False)


async def blob_get_async(pathname: str) -> bytes | None:
    """Returns None (rather than raising) whenever the blob simply doesn't exist --
    callers use this purely to check "is there a signature saved for this company/id",
    which is a normal, expected case, not an error."""
    from vercel.blob import AsyncBlobClient

    client = AsyncBlobClient()
    try:
        result = await client.get(pathname, access="public")
    except Exception:
        return None
    if result is None:
        return None
    chunks = []
    async for chunk in result.stream:
        chunks.append(chunk)
    return b"".join(chunks)


async def blob_delete_async(pathname: str) -> None:
    from vercel.blob import AsyncBlobClient

    client = AsyncBlobClient()
    try:
        await client.delete(pathname)
    except Exception:
        pass


def blob_put(pathname: str, data: bytes) -> None:
    asyncio.run(blob_put_async(pathname, data))


def blob_get(pathname: str) -> bytes | None:
    return asyncio.run(blob_get_async(pathname))


def blob_delete(pathname: str) -> None:
    asyncio.run(blob_delete_async(pathname))
