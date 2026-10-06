cmake_minimum_required(VERSION 3.10)
project(marisa VERSION 0.2.6)

# Official repo: https://github.com/s-yata/marisa-trie
# CMake script borrowed from https://github.com/Karry/osmscout-sailfish/blob/master/CMakeLists.txt

if(UNIX)
  set (CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -Weffc++ -Wextra -Wconversion")
endif()

set(MARISA_ROOT ${CMAKE_CURRENT_SOURCE_DIR})

set(MARISA_INCLUDE_DIRS
    ${MARISA_ROOT}/lib
    ${MARISA_ROOT}/include)

set(MARISA_SRC
    ${MARISA_ROOT}/lib/marisa/grimoire/io/reader.cc
    ${MARISA_ROOT}/lib/marisa/grimoire/io/mapper.cc
    ${MARISA_ROOT}/lib/marisa/grimoire/io/writer.cc
    ${MARISA_ROOT}/lib/marisa/grimoire/trie/tail.cc
    ${MARISA_ROOT}/lib/marisa/grimoire/trie/louds-trie.cc
    ${MARISA_ROOT}/lib/marisa/grimoire/vector/bit-vector.cc
    ${MARISA_ROOT}/lib/marisa/trie.cc
    ${MARISA_ROOT}/lib/marisa/agent.cc
    ${MARISA_ROOT}/lib/marisa/keyset.cc
)

include_directories(
  ${MARISA_INCLUDE_DIRS}
)

include(CMakePackageConfigHelpers)
include(GNUInstallDirs)

configure_package_config_file(
  marisa-config.cmake.in
  ${CMAKE_CURRENT_BINARY_DIR}/marisa-config.cmake
  INSTALL_DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/marisa
  PATH_VARS CMAKE_INSTALL_INCLUDEDIR
)

write_basic_package_version_file(
  ${CMAKE_CURRENT_BINARY_DIR}/marisa-version.cmake
  COMPATIBILITY SameMajorVersion
)

add_library(marisa
    ${MARISA_SRC}
)

install(TARGETS marisa
        EXPORT marisa-targets
        RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
        LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
        ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR})

set(include_headers ${MARISA_ROOT}/include/marisa.h)
file(GLOB libmarisa_include_headers ${MARISA_ROOT}/include/marisa/*.h)
install(FILES ${include_headers} DESTINATION include)
install(FILES ${libmarisa_include_headers} DESTINATION include/marisa)
target_include_directories(marisa PUBLIC
  $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
  $<INSTALL_INTERFACE:${CMAKE_INSTALL_INCLUDEDIR}>
)

install(
  EXPORT marisa-targets
  DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/marisa
)

install(
  FILES
    ${CMAKE_CURRENT_BINARY_DIR}/marisa-config.cmake
    ${CMAKE_CURRENT_BINARY_DIR}/marisa-version.cmake
  DESTINATION
    ${CMAKE_INSTALL_LIBDIR}/cmake/marisa)
