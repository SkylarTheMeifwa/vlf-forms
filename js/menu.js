document.addEventListener("DOMContentLoaded", function () {

    fetch("../assets/menu.html")
        .then(response => response.text())
        .then(data => {

            document.getElementById("menu-container").innerHTML = data;

            const currentPage = window.location.pathname;

            document.querySelectorAll("#menu a").forEach(link => {

                const linkURL =
                    new URL(link.href, window.location.href);

                if (linkURL.pathname === currentPage) {
                    link.classList.add("active");
                }

            });


            // =====================================================
            // DROPBOX TEMPLATE ACCESS
            // =====================================================

            document.querySelectorAll(
                "#menu a[data-template]"
            ).forEach(link => {

                link.addEventListener(
                    "click",
                    async function (event) {

                        event.preventDefault();

                        const templateName =
                            this.dataset.template;

                        const formURL =
                            this.href;

                        await openFormIfAccessible(
                            templateName,
                            formURL
                        );

                    }
                );

            });

        });

});