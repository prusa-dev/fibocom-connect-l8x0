mkdir -Force ./dist | Out-Null

$git_hash = $(git rev-parse --short HEAD)

Compress-Archive -Force -CompressionLevel Optimal `
    -DestinationPath "./dist/fibocom-connect.$git_hash.zip" `
    -Path './*.cmd', './scripts', './drivers', './screenshot', './*.md'
