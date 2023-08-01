var menuTimer = null;

function showMenu(mid, loginid, e)
{
    var divMenu = getElement("divMenu");
    if(divMenu == null)
    {
        return;
    }
    divMenu.innerHTML = menuHTML.replace("__id__", mid).replace("__id__", mid).replace("__loginid__", loginid).replace("__loginid__", loginid).replace("__loginid__", loginid);
    divMenu.style.display = "";
    
    divMenu.style.left = currentX + "px";
    divMenu.style.top = currentY + "px";
    
    menuHideAction();
}

function menuHideAction()
{
    menuTimer = setInterval("menuHide()", 500);
}

function menuShowAction()
{
    clearInterval(menuTimer);
}

function menuHide()
{
    var divMenu = getElement("divMenu");
    if(divMenu == null)
    {
        return;
    }
    
    divMenu.style.display = "none";    
    clearInterval(menuTimer);
}

