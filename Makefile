.PHONY: docs docs-serve docs-build docs-clean test lint format
.ONESHELL:
SHELL := /bin/bash

PYTHON ?= python3

# Documentation (MkDocs + Material theme)
docs: docs-serve

docs-serve:
	uvx --with mkdocs-material mkdocs serve

docs-build:
	uvx --with mkdocs-material mkdocs build

docs-clean:
	rm -rf site/

# Testing
test:
	$(PYTHON) -m pytest tests/ -v

# Linting & formatting
lint:
	uvx ruff check src/ tests/ examples/

format:
	uvx ruff format src/ tests/ examples/
