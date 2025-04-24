from groq import AsyncGroq
import asyncio
import os
from dotenv import load_dotenv

from main import get_knn_chunks
from pathlib import Path
import json
from pprint import pprint

load_dotenv()


def get_context(query: str):
    root_dir = Path(__file__).resolve(strict=True).parent

    source_file_dir = str(root_dir / "source_files")
    vector_db_dir = str(root_dir / "vector_dbs")

    # pass threshold(distance) < 0.6 for most relevent results(D < 0.8 => moderat results)cosine similarity => 1 - distance / 2
    context = get_knn_chunks(source_file_dir, vector_db_dir, query, threshold=1, k=5)
    pprint(context)
    context = "\n\n".join([str(i.get("text", None)) for i in context])
    return context


async def main(query: str, context: str = ""):
    client = AsyncGroq(api_key=os.getenv("GROQ_API_KEY"))

    chat_completion = await client.chat.completions.create(
        # response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": f"""You are Microsoft Excel expert AI assistant.
                You can help users with Excel formulas, functions, and data analysis.
                You can also help users with Excel VBA and macros.
                You can also help users with Excel Power Query and Power Pivot.
                You can also help users with Excel Power BI and Power Automate.
                You can also help users with Excel Power Apps and Power Automate Desktop.
                You can also help users with Excel Power Query M language and DAX.
                Your response must explain the solution step by step so user can understand it easily.""",
            },
            {
                "role": "user",
                "content": f"""{query}\n\nYou can use the following contextual data
                if it is correct and relevant to the query.
                DO NOT quote the context data in your response.
                Ignore the context data if it is irrelevent or empty.,
                context: '{context}'.""",
            },
        ],
        temperature=0.3,
        max_completion_tokens=1024,
        # model="llama-3.3-70b-versatile",
        model="llama-3.1-8b-instant",
    )
    # print(chat_completion.choices[0].message.content)
    return chat_completion.choices[0].message.content


async def test():
    query: str = (
        """
        How to flip columns into rows in excel?
        """
    )
    context = get_context(query)
    res = await main(query, context)
    print(res)
    with open("testResults8b.json", "a") as f:
        f.write(json.dumps({f"(with official docs)\tquery({query})\ncontext({context})": res}, indent=4))


if __name__ == "__main__":
    asyncio.run(test())
