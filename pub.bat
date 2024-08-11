G:
cd G:\projects\TimeIngest\TimeIngest\
del G:\projects\TimeIngest\TimeIngest\mypythonscript.py
copy G:\projects\TimeIngest\TimeIngest\bin\Debug\net8.0\mypythonscript.py .
dotnet publish timeingest.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
copy G:\projects\TimeIngest\TimeIngest\mypythonscript.py G:\projects\TimeIngest\TimeIngest\bin\Release\net8.0\win-x64\publish\mypythonscript.py
