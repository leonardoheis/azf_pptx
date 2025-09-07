# azf_pptx

Development setup with Pipenv

Install dependencies (development):

```bash
pipenv install --dev
```

Run tests inside the Pipenv shell:

```bash
pipenv run pytest -q
```

If you want only production deps (no dev):

```bash
pipenv install --deploy --ignore-pipfile
```