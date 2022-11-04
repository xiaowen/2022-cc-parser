Read [this blog post](https://xiaowenx.medium.com/parsing-credit-card-statements-using-machine-learning-on-google-cloud-and-azure-65df6bf39ee8) for info on this repo.

The main code is in `main.py`.  To run it, configure the environment as follows:
1. (Optional) set up Python [venv](https://docs.python.org/3/library/venv.html).
2. Install Python libraries: `pip3 install google-api-python-client google-auth-httplib2 google-auth-oauthlib google-cloud-documentai azure azure-ai-formrecognizer`
3. Configure the environment to use a Google Cloud service account.  If using a key file, then set `GOOGLE_APPLICATION_CREDENTIALS`.  If using vscode and friends, then use `env` to set it.
4. Put your Azure API key into the `AZURE_COGNITIVE_SERVICES_KEY` enviroment variable.