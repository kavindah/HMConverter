SIZE CHART EXTRACTOR - free web app
================================================

Goal: a public link where anyone can upload a tech-pack PDF and download
the Excel size chart. Hosted free on Streamlit Community Cloud.

FILES
  streamlit_app.py    the web page (upload -> convert -> download)
  sizechart_app.py    the extraction engine (shared with the desktop app)
  requirements.txt    the Python packages the host needs to install

--------------------------------------------------------------------
TRY IT ON YOUR OWN COMPUTER FIRST (optional)
--------------------------------------------------------------------
  pip install -r requirements.txt
  streamlit run streamlit_app.py
A browser tab opens at http://localhost:8501.

--------------------------------------------------------------------
PUT IT ONLINE FOR FREE (Streamlit Community Cloud)
--------------------------------------------------------------------
1. Make a free GitHub account:        https://github.com
2. Create a new repository, e.g. "size-chart-app".
3. Upload these three files into it:
       streamlit_app.py
       sizechart_app.py
       requirements.txt
   (GitHub website: "Add file" -> "Upload files" -> drag them in -> Commit.)
4. Go to https://share.streamlit.io  and sign in with GitHub.
5. Click "Create app" -> "Deploy a public app from GitHub".
       Repository:      your-username/size-chart-app
       Branch:          main
       Main file path:  streamlit_app.py
6. Click Deploy. After a minute you get a public URL like
       https://size-chart-app.streamlit.app
   Share that link - anyone can open it and convert their own PDFs.

Updating later: just edit the files on GitHub; the app redeploys itself.

--------------------------------------------------------------------
ALTERNATIVE FREE HOST: Hugging Face Spaces
--------------------------------------------------------------------
1. Free account at https://huggingface.co
2. Create a new "Space" -> choose the "Streamlit" SDK.
3. Upload the same three files. It builds and serves automatically,
   giving a public URL like https://huggingface.co/spaces/you/size-chart.

--------------------------------------------------------------------
PRIVACY / CONFIDENTIALITY
--------------------------------------------------------------------
- The code in the repo contains no tech-pack data; users upload their own
  PDFs at runtime and the app processes them in memory to build the Excel.
- A public app means anyone with the link can use it and uploads go to the
  host's servers. If your tech packs are confidential, either keep the repo
  private and use Streamlit's per-app viewer allow-list, or host it only on
  an internal/company machine instead of a public free host.
