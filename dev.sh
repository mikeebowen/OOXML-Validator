#!/usr/bin/env bash
set -e 

build_envs=(
    linux-x64
    linux-arm64
    osx-x64
    osx-arm64
    win-x64
)

help () {
    cat <<EOF
$0 <command> [build_env]

  help                 show this help message
  docker               docker container for development
  build <build_env>    build
  envs                 show build envs
  run <build_env>      run the command line

EOF
}

if [ "$#" -lt 1 ]; then
    help
fi

invalid_env () {
    echo "Error: '$1' must be one of '${build_envs[@]}'"
    help
}

command=$1;

build () {
    if [ "$#" -ne 1 ]; then
        echo "Illegal number of parameters"
        help
        exit 1
    fi
    build_env=$1

    if ! [[ ${build_envs[@]} =~ $build_env ]]; then
        invalid_env $build_env
        exit 1
    fi

    case "$build_env" in
        "linux-x64")
            dotnet publish --framework net8.0 -c Release -r linux-x64 -p:IncludeNativeLibrariesForSelfExtract=true -p:InvariantGlobalization=true OOXMLValidator.sln
            ;;
        "linux-arm64")
            dotnet publish --framework net8.0 -c Release -r linux-arm64 -p:IncludeNativeLibrariesForSelfExtract=true -p:InvariantGlobalization=true OOXMLValidator.sln
            ;;
        "osx-x64")
            dotnet publish --framework net8.0 -c Release -r osx-x64 OOXMLValidator.sln
            ;;
        "osx-arm64")
            dotnet publish --framework net8.0 -c Release -r osx-arm64 OOXMLValidator.sln
            ;;
        "win-x64")
            dotnet publish --framework net8.0 -c Release -r win-x64 OOXMLValidator.sln
            ;;
    esac

    if ! [[ "$build_env" =~ "win-" ]]; then
        chmod +x "./OOXMLValidatorCLI/bin/Release/net8.0/${build_env}/publish/OOXMLValidatorCLI"
    fi
}

run () {
    if [ "$#" -lt 1 ]; then
        echo "Illegal number of parameters"
        help
        exit 1
    fi
    build_env=$1

    if ! [[ ${build_envs[@]} =~ $build_env ]]; then
        invalid_env "$build_env"
        exit 1
    fi

    if [[ "$build_env" =~ "win-" ]]; then
        ext=".exe"
    fi

    "./OOXMLValidatorCLI/bin/Release/net8.0/${build_env}/publish/OOXMLValidatorCLI${ext}" ${*:2}
    exit 0
}

test () {
    if [ "$#" -lt 1 ]; then
        echo "Illegal number of parameters"
        help
        exit 1
    fi
    build_env=$1

    if ! [[ ${build_envs[@]} =~ $build_env ]]; then
        invalid_env "$build_env"
        exit 1
    fi

    if [[ "$build_env" =~ "win-" ]]; then
        ext=".exe"
    fi

    shell_cmd="./${CI_SHELL_OVERRIDE:-"OOXMLValidatorCLI/bin/Release/net8.0/${build_env}/publish"}/OOXMLValidatorCLI${ext}"
    echo $shell_cmd
    chmod +x "${shell_cmd}"
    
    output="$($shell_cmd)" 
    if [[ "$output" == "Value cannot be null." ]]; then
        echo "success"
        exit 0
    else
        echo "error:"
        echo "$output"
        exit 2 
    fi
    
}

envs () {
    echo "${build_envs[@]}"
}

docker () {
    if [[ "$DOCKER_RUNNING" == "true" ]]; then
        echo "Error: Already running inside container"
        help
        exit 1
    fi

    cleanup() {
        rm .build-files/Dockerfile || true
        rm .build-files/compose.yaml || true
        rmdir .build-files || true
    }
    trap cleanup EXIT

    mkdir .build-files
    cat << EOF > .build-files/Dockerfile
FROM ubuntu
ENV DEBIAN_FRONTEND noninteractive

RUN apt-get update && apt-get install -y dotnet-sdk-8.0

WORKDIR /code
EOF

    cat << EOF > .build-files/compose.yaml
services:
  dev:
    build:
      dockerfile: ./Dockerfile
    environment:
      - DOCKER_RUNNING=true
    volumes:
      - ../:/code
EOF

    docker-compose -f .build-files/compose.yaml run dev bash
}

case "$command" in
    "docker")
        docker
        ;;
    "build")
        build $2
        ;;
    "run")
        run ${*:2}
        ;;
    "envs")
        envs
        ;;
    "help")
        help
        ;;
    "test")
        test $2
        ;;
esac
