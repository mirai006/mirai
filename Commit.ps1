

$DateTime = (Get-Date -Format "yyyy/MM/dd HH:mm") + " (JST)"

echo "cscript vbac.wsf decombine`n"
cscript vbac.wsf decombine

echo "git add`n"
git add .

echo "git commit`n"
git commit -m "$args Commited at $DateTime"

echo "git push master`n"
git push origin master
echo "git push origin`n"
git push origin
