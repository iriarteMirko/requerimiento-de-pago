import warnings
from app import App


def main():
    warnings.filterwarnings("ignore")
    app = App()
    app.crear_app()


if __name__ == "__main__":
    main()