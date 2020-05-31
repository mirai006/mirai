
Param($Reson)
$DateTime = (Get-Date -Format "yyyy/MM/dd HH:mm") + " (JST)"

cscript vbac.wsf decombine

git add ./*.*
git add ./bin/*.*
git add ./src

git commit -m "$Reson Commited at $DateTime"

git push origin master
git push origin
