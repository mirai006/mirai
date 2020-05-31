
Param($Reson)

$DateTime = (Get-Date -Format "yyyy/MM/dd HH:mm") + " (JST)"

git add .

git commit -m "$Reson Commit at $DateTime"

git push origin master
git push origin
