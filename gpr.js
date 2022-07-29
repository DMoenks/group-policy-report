function closeAll()
{
    var nodes = document.getElementsByClassName("comparison");
    for (var i = 0; i < nodes.length; i++)
    {
        nodes[i].style.display = "none";
    }
}

function openSelected(selector)
{
    closeAll();
    var selection = document.getElementById(selector);
    document.getElementById(selection.options[selection.selectedIndex].value).style.display = "block";
}

function unbindScroll(side)
{
    document.getElementById(side).onscroll = null;
}

function scroll(master, slave)
{
    var tmpmaster = document.getElementById(master);
    var tmpslave = document.getElementById(slave);
    var percentageX = tmpmaster.scrollLeft / (tmpmaster.scrollWidth - tmpmaster.offsetWidth);
    tmpslave.scrollLeft = (tmpslave.scrollWidth - tmpslave.offsetWidth) * percentageX;
    var percentageY = tmpmaster.scrollTop / (tmpmaster.scrollHeight - tmpmaster.offsetHeight);
    tmpslave.scrollTop = (tmpslave.scrollHeight - tmpslave.offsetHeight) * percentageY;
}

function bindScroll(side)
{
    document.getElementById(side).onscroll = function onscroll(event) { scroll(side); };
}

window.onload = closeAll();
