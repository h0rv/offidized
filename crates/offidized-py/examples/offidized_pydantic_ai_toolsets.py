"""Minimal PydanticAI toolset setup for offidized.

Run from the offidized-py crate directory:
    cd crates/offidized-py && uv run python examples/offidized_pydantic_ai_toolsets.py

Requires:
    uv sync --extra pydantic-ai
or:
    pip install "offidized[pydantic-ai]"
"""

from pydantic_ai import Agent

from offidized.pydantic_ai import (
    all_toolsets,
    compose_toolsets,
    docx_toolset,
    pptx_toolset,
    xlsx_toolset,
)


def build_agents() -> tuple[Agent, Agent, Agent]:
    """Construct agents with separate, composed, and all-in-one toolsets."""
    format_agent = Agent(
        "openai:gpt-5",
        toolsets=[
            xlsx_toolset(),
            docx_toolset(),
            pptx_toolset(),
        ],
    )

    combined_agent = Agent(
        "openai:gpt-5",
        toolsets=[all_toolsets()],
    )

    composed_agent = Agent(
        "openai:gpt-5",
        toolsets=[compose_toolsets(xlsx_toolset(), docx_toolset())],
    )

    return format_agent, combined_agent, composed_agent


if __name__ == "__main__":
    agent, combined, composed = build_agents()
    print(f"Built format agent with {len(agent.toolsets)} toolsets")
    print(f"Built combined agent with {len(combined.toolsets)} toolset")
    print(f"Built composed agent with {len(composed.toolsets)} toolset")
