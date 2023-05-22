mkdir -Force ./dist | Out-Null

Compress-Archive -Force -CompressionLevel Optimal `
    -DestinationPath ./dist/fibocom-connect.zip `
    -Path './*.cmd', './scripts', './screenshot', './*.md'
