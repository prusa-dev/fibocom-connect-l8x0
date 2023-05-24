mkdir -Force ./dist | Out-Null

$git_hash = $(git describe --tags)

Compress-Archive -Force -CompressionLevel Optimal `
    -DestinationPath "./dist/fibocom-connect.$git_hash.zip" `
    -Path './*.cmd', './scripts', './screenshot', './*.md'

Compress-Archive -Force -CompressionLevel Optimal `
    -DestinationPath "./dist/drivers_l860.zip" `
    -Path './drivers'
