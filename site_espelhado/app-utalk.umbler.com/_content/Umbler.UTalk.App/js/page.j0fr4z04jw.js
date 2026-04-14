function clearSelection() {
    if (document.selection && document.selection.empty)
        document.selection.empty();
    else if (window.getSelection) window.getSelection().removeAllRanges();
}

/**
 * @param {Element} element
 * @param {boolean} toTop
 */
function scrollElementToBottom(element) {
    element.scrollTop = element.scrollHeight;
}

/** @param {Element} multiFilterRoot */
function multiFilterScrollActiveIntoView(multiFilterRoot) {
    var activeItem = multiFilterRoot.getElementsByClassName(
        "dropdown-item active"
    )[0];

    if (activeItem) {
        var block = "nearest";
        if (activeItem.classList.contains("first-item")) block = "end";
        else if (activeItem.classList.contains("last-item")) block = "start";

        activeItem.scrollIntoView({
            behavior: "smooth",
            block: block,
            inline: "nearest",
        });
    }
}

function getHeight(element) {
    return element.clientHeight;
}

async function uninstallServiceWorkers() {
    for (var sw of await navigator.serviceWorker.getRegistrations())
        await sw.unregister();
}

function adjustTextAreaHeight(id) {
    const textarea = document.getElementById(id);
    if (textarea) {
        textarea.style.height = 'auto';
        textarea.style.height = (textarea.scrollHeight) + 'px';
    }
}

function animateElementBounceIn(elementId) {

    setTimeout(() => {
    let div = document.getElementById(elementId);
    if (!div) return;

    div.removeAttribute("aria-hidden");
        div.classList.remove("animate__bounceIn", "highlight");

    setTimeout(() => {
        div.classList.add("animate__bounceIn", "highlight");
    }, 10);

        
  

    div.addEventListener("animationend", () => {
        div.classList.remove("animate__bounceIn", "highlight");

        setTimeout(() => {
            div.classList.add("highlight");
        }, 100);

    }, { once: true });
    }, 300);
};

function initializeHotReloadThemeKeeper() {
    var htmlElement = document.body.parentElement;
    var lastTheme;
    var observer = new MutationObserver((mutations) => {
        if (
            !mutations.some(
                (m) => m.type == "attributes" && m.attributeName == "data-bs-theme"
            )
        )
            return;

        var currentTheme = htmlElement.getAttribute("data-bs-theme");
        if (currentTheme === "dark") lastTheme = "dark";
        else if (currentTheme === "light-emerald") lastTheme = "light-emerald";
        else if (currentTheme === "dark-emerald") lastTheme = "dark-emerald";
        else if (currentTheme === "light") lastTheme = "light";
        else htmlElement.setAttribute("data-bs-theme", lastTheme || "dark");
    });

    observer.observe(htmlElement, {
        attributes: true,
        attributeFilter: ["data-bs-theme"],
    });
}

initializeHotReloadThemeKeeper();
