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

    function setTextAnimation(delay, duration, strokeWidth, timingFunction, strokeColor, repeat) {
        let paths = document.querySelectorAll(".termtxt svg path");
        let mode = repeat ? 'infinite' : 'forwards'
        for (let i = 0; i < paths.length; i++) {
            const path = paths[i];
            const length = path.getTotalLength();
            path.style["stroke-dashoffset"] = `${length}px`;
            path.style["stroke-dasharray"] = `${length}px`;
            path.style["stroke-width"] = `${strokeWidth}px`;
            path.style["stroke"] = `${strokeColor}`;
            path.style["animation"] = `${duration}s svg-text-anim ${mode} ${timingFunction}`;
            path.style["animation-delay"] = `${i * delay}s`;
        }
    }

    setTextAnimation(0.4, 8, 2, 'ease', '#423658', true);

    // ─── Search commands ───────────────────────────────────────────────────────
    const searchInput = document.querySelector(".terminal-head .search input");
    const blocks = termBody.querySelectorAll(".block");

    const getBlockText = (block) => {
        const comment = block.querySelector(".comment");
        const command = block.querySelector(".command");
        const parts = [
            comment?.textContent?.trim() ?? "",
            command?.textContent?.trim() ?? "",
        ];
        return parts.join(" ").toLowerCase();
    };

    const filterBlocks = (query) => {
        const q = query.trim().toLowerCase();
        blocks.forEach((block) => {
            const text = getBlockText(block);
            const match = !q || text.includes(q);
            block.classList.toggle("is-hidden", !match);
        });
    };

    searchInput.addEventListener("input", (e) => filterBlocks(e.target.value));
})()