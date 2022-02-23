if [ -z $1 ]; then
swift test --enable-test-discovery
else
swift test --enable-test-discovery --filter $1
fi
