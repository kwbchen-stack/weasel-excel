rem Build settings for Weasel (Excel dictionary edition v4), Visual Studio 2022

rem REQUIRED: path to Boost 1.84.0 source directory
if not defined BOOST_ROOT set BOOST_ROOT=C:\Libraries\boost_1_84_0

rem Visual Studio 2022 toolset
set BJAM_TOOLSET=msvc-14.3
set CMAKE_GENERATOR="Visual Studio 17 2022"
set PLATFORM_TOOLSET=v143

rem For Visual Studio 2019 use these three lines instead:
rem set BJAM_TOOLSET=msvc-14.2
rem set CMAKE_GENERATOR="Visual Studio 16 2019"
rem set PLATFORM_TOOLSET=v142
