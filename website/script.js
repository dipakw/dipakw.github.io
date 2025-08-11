(function () {
    const termBody = document.querySelector(".terminal-body");

    const copyToClipboard = (text) => {
        try {
            const textArea = document.createElement("textarea");
            textArea.value = text;

            // Avoid scrolling to bottom
            textArea.style.position = "fixed";
            textArea.style.top = "-9999px";
            textArea.style.left = "-9999px";

            document.body.appendChild(textArea);
            textArea.focus();
            textArea.select();

            const successful = document.execCommand('copy');
            document.body.removeChild(textArea);
            return successful;
        } catch (err) {
            console.error('Fallback: Unable to copy', err);
            return false;
        }
    }

    termBody.addEventListener("click", (e) => {
        let line = null;

        if (e.target.classList.contains("command")) {
            line = e.target;
        } else {
            line = e.target.closest(".command");
        }

        if (!line) {
            return;
        };

        const command = Array.from(line.children).map(child => child.textContent).join(" ");

        if (copyToClipboard(command)) {
            var toast = Toastify({
                text: "Copied to clipboard",
                duration: 1500,
                close: false,
                gravity: "bottom",
                position: "center",
                stopOnFocus: false,
                style: {
                    background: "#423658",
                    fontSize: "14px",
                    fontWeight: "normal",
                    borderRadius: "8px",
                    padding: "8px 16px",
                },
            });

            toast.showToast();
        }
    });
})()