cmake_minimum_required(VERSION 3.10)

project(VPR-Convertor-New)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Qt6 REQUIRED COMPONENTS Widgets Core Gui AxContainer)


set(SOURCES 
"${CMAKE_SOURCE_DIR}/source/Table.cpp"
"${CMAKE_SOURCE_DIR}/source/main.cpp"
"${CMAKE_SOURCE_DIR}/source/Table.h"
) # Добавьте сюда ваши файлы

qt6_wrap_cpp(MOC_SOURCES ${SOURCES})

add_executable(VPR-Convertor-New-App ${SOURCES} ${MOC_SOURCES})

target_link_libraries(VPR-Convertor-New-App PRIVATE Qt6::Widgets Qt6::Core Qt6::Gui Qt::AxContainer)