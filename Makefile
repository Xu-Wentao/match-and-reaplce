.PHONY: build build-mac build-win-match build-win-replace

build-mac:
	GOOS=darwin go build -o bin/app-drawin

build-win: build-win-match build-win-replace

build-win-match:
	GOOS=windows GOARCH=amd64 go build  -o bin/app-匹配-amd64.exe

build-win-replace:
	GOOS=windows GOARCH=amd64 go build -o bin/app-替换-amd64.exe
