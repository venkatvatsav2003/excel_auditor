.PHONY: all install audit test clean docker

VERSION ?= 2.0.0

all: install

install:
	pip install pyyaml

audit:
	./audit.sh $(FILES)

test:
	python3 -m pytest tests/ -v

clean:
	rm -rf reports/ *.pyc __pycache__ .pytest_cache

docker:
	docker build -t data-auditor:$(VERSION) .
