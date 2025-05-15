@echo off

rmdir /s /q build
mkdir build
cd build

cmake -DCMAKE_PREFIX_PATH="C:\Qt\6.2.4\msvc2019_64" ..


cmake --build .  --config Release
cd ..