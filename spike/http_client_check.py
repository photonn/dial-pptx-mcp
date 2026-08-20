#!/usr/bin/env python
"""Quick streamable-http transport check: list tools and exercise a create/save cycle."""
import asyncio
import json
import sys

from mcp import ClientSession
from mcp.client.streamable_http import streamablehttp_client

URL = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8123/mcp"


async def main():
    async with streamablehttp_client(URL) as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = await session.list_tools()
            print(f"tools: {len(tools.tools)}")
            r = await session.call_tool("create_presentation", {})
            print("create:", r.content[0].text[:200])
            # Every call needs the handle explicitly: the server has no
            # "current presentation" (see state.py).
            pres_id = json.loads(r.content[0].text)["presentation_id"]
            r = await session.call_tool("add_slide", {
                "layout_index": 0, "title": "HTTP check",
                "presentation_id": pres_id,
            })
            print("add_slide:", r.content[0].text[:200])


asyncio.run(main())
