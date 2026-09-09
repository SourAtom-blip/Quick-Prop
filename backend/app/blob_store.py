"""Optional persistent storage backend for signature images via Vercel Blob.

Local disk (SIGNATURES_DIR/TEMP_SIGNATURES_DIR in renderer.py) works fine on a normal
server with a persistent filesystem, but breaks on Vercel: every serverless invocation
gets a fresh, isolated /tmp, so a signature uploaded in one request is gone by the next.
Vercel automatically sets BLOB_READ_WRITE_TOKEN when Blob storage is enabled on the
project -- its presence is what we use to decide which backend to use, so local dev
(no token set) keeps working exactly as before, untouched.
"""
import asyncio
import os

BLOB_ENABLED = bool(os.environ.get("BLOB_READ_WRITE_TOKEN"))


def blob_put(pathname: str, data: bytes) -> None:
    from vercel.blob import AsyncBlobClient

    async def _put():
        client = AsyncBlobClient()
        await client.put(pathname, data, access="public", add_random_suffix=False)

    asyncio.run(_put())


def blob_get(pathname: str) -> bytes | None:
    """Returns None (rather than raising) whenever the blob simply doesn't exist --
    callers use this purely to check "is there a signature saved for this company/id",
    which is a normal, expected case, not an error."""
    from vercel.blob import AsyncBlobClient

    async def _get():
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

    return asyncio.run(_get())


def blob_delete(pathname: str) -> None:
    from vercel.blob import AsyncBlobClient

    async def _del():
        client = AsyncBlobClient()
        try:
            await client.delete(pathname)
        except Exception:
            pass

    asyncio.run(_del())
