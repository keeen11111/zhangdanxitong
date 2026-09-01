import uvicorn


def main() -> None:
    uvicorn.run("main:app", host="127.0.0.1", port=8002, app_dir="api/py", log_level="warning")


if __name__ == "__main__":
    main()

