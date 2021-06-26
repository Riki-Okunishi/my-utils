# my finctions

# wrapper for ns3
ns3() {
    if [ $1 = 'run' -a $# -ge 2 ]; then
        shift
        ./waf --run "$@"
    else
        ./waf "$@"
    fi
}

# wrapper for ns3 in sudo
suns3() {
    if [ $1 = 'run' -a $# -ge 2 ]; then
        shift
        sudo -S ./waf --run "$@"
    else
        sudo -S ./waf "$@"
    fi
}