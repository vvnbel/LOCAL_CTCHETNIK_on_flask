from loguru import logger
from app import app

if __name__ == '__main__':
    logger.add("runtime.log", rotation="100 MB")
    app.run(host="0.0.0.0", port="5050", debug=True)