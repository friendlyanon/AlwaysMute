cmake_minimum_required(VERSION 3.20)
project(AlwaysMute CXX)
add_executable(AlwaysMute WIN32 main.cpp)
target_compile_features(AlwaysMute PRIVATE cxx_std_23)
target_compile_definitions(
    AlwaysMute PRIVATE
    WIN32_LEAN_AND_MEAN=1
    NOMINMAX=1
    NTDDI_VERSION=0x0A000007
    _WIN32_WINNT=0x0A00
    UNICODE=1
    _UNICODE=1
)
install(TARGETS AlwaysMute)
