import os

# gunicorn configuration
workers = 4
worker_class = 'sync'
worker_connections = 1000
timeout = 30
keepalive = 2

# Get port from environment variable or use default
bind = f"0.0.0.0:{os.environ.get('PORT', '5000')}"
