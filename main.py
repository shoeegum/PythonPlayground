
from app import app

if __name__ == "__main__":
    # Use gunicorn for production
    import gunicorn.app.base

    class StandaloneApplication(gunicorn.app.base.BaseApplication):
        def __init__(self, app, options=None):
            self.options = options or {}
            self.application = app
            super().__init__()

        def load_config(self):
            for key, value in self.options.items():
                self.cfg.set(key, value)

        def load(self):
            return self.application

    options = {
        'bind': '0.0.0.0:5000',
        'workers': 4,
        'worker_class': 'sync',
        'timeout': 120
    }

    StandaloneApplication(app, options).run()
