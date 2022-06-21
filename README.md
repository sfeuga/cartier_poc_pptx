# POC PPTX Generator for Cartier

## Dependencies

- Python 3.10
- [Fake Cartier API](https://github.com/sfeuga/cartier_mock_api) running (locally or remotely)

Optional:
- direnv

## Setting up (and running) a local environment

### Local install
  1. Create a virtual environment: `python3 -m venv venv`
  2. Activate the virtual environment: `source venv/bin/activate`
  3. Install python dependencies: `pip install -e .`
  4. Set up your `.envrc` by copying `.envrc.dist` to `.envrc`
     - ⚠️ If __Fake Cartier API__ is running remotely, change the url in your `.envrc`
  5. Source your Env. Var.: `source .envrc` or allow it with direnv
  6. Run `python3 app.py`
