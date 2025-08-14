# gunicorn configuration
workers = 4
worker_class = 'sync'
worker_connections = 1000
timeout = 30
keepalive = 2
bind = '0.0.0.0:' + str(int('$PORT')) if '$PORT' in os.environ else '0.0.0.0:5000'
