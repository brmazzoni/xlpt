

build: Dockerfile
		docker build -t xlpt .

instanciate:
		docker run --rm -v $(shell pwd):/xlpt xlpt instanciate

generate:
		docker run --rm -v $(shell pwd):/xlpt xlpt generate

clean:
		@echo "Deleting following files and directories:"
		@rm -vrf *.xlsx xlpt/*.xlsx xlpt/__pycache__
