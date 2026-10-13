<div align="center">
  <img src="./banner.svg" alt="azure-enum" width="800">
</div>

# azure-enum

> Red team security tool — authorized use only.

```bash
python azure_enum.py --help
```

> **Authorized security testing only.**

Inventory requests follow Microsoft Graph continuation links so results are not
silently limited to the first page. Continuation URLs are restricted to
`https://graph.microsoft.com/v1.0/`, and each collection is capped at 20 pages;
the tool reports an error rather than returning a partial collection when that
limit is exceeded. All validation and tests can run offline with mocked Graph
responses.

**Omar Khalid** — [omareldemery.com](https://omareldemery.com)
