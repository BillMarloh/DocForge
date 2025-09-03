#!/bin/sh
set -e
IMAGE=${IMAGE:-file-converter}
PORT=${PORT:-3456}
TOKEN=${FILE_TOOL_TOKEN:-}
docker build -t $IMAGE .
docker rm -f $IMAGE || true
docker run -d --name $IMAGE -p $PORT:3456 \ 
  -e FILE_TOOL_TOKEN=$TOKEN $IMAGE
echo "Running at http://localhost:$PORT"


