for ($i=1; $i -le 50; $i++) {
    dotnet test
    if($?)
    {
        echo "command succeeded"
    }
    else
    {
        echo "command failed"
        break
    }
}
