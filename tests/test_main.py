from zoemzone.main import run


def get_greeting() -> str:
    return "Hello from ZoemZone!"


def test_run_prints_greeting(capsys):
    run()
    captured = capsys.readouterr()
    greeting = captured.out.strip()
    print(greeting)
    assert greeting == get_greeting()


if __name__ == "__main__":
    print(get_greeting())
