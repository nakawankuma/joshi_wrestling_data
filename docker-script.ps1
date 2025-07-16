docker run -it --rm  --name container-joshi `
  -v "C:\src\docker-claudecode\2女子プロレスラーの一覧:/src" `
  -v "C:\src\docker-claudecode\home:/root" `
  image-joshi  /bin/bash -c "python ./xlsx_to_html_converter_all_in_one.py"


