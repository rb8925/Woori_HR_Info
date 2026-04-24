#!/usr/bin/env bash
# Nix 환경에서 numpy/pandas에 필요한 시스템 라이브러리 경로를 등록하고 실행한다.
ZLIB="/nix/store/0zv8lswa9k122sixl00zjb1g1r49bs0i-zlib-1.3/lib"
STDCPP="/nix/store/55byk2fn6548ni8ibgd2dyzpmk4z180w-gcc-12.2.0-lib/lib"
export LD_LIBRARY_PATH="$ZLIB:$STDCPP:$LD_LIBRARY_PATH"

source "$(dirname "$0")/.venv/bin/activate"
exec "$@"
