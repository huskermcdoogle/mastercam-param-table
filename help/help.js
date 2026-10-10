/* Parameter Table Tool - user manual: the menu button and the search.
   No internet, no libraries. The search index is search.js (made by tools/make_help.py),
   loaded as a plain script so it works from a folder (file://) in any browser. */
(function () {
    "use strict";
    document.documentElement.className += " js";

    // Words people type that the pages say another way.
    var SAME = {
        faster: ["time", "cycle", "slowest", "target"], quicker: ["time", "cycle"], slow: ["slowest", "time"],
        sfm: ["speed", "css", "surface"], smm: ["speed", "css", "surface"], rpm: ["speed", "spindle"],
        ipr: ["rev", "feed"], ipm: ["min", "feed"], feedrate: ["feed"], rate: ["feed", "mrr"],
        flip: ["flips", "insert"], rotate: ["flips", "insert"], index: ["flips", "insert"],
        insert: ["inserts", "flips", "edge"], life: ["edge", "inspection"], wear: ["edge", "inspection"],
        cost: ["cost", "inserts"], money: ["cost"], price: ["cost"],
        mistake: ["undo", "back"], oops: ["undo"], revert: ["back", "undo"], restore: ["undo", "scenario"],
        print: ["report", "print"], pdf: ["print", "report"], boss: ["report"],
        coolant: ["coolant", "flood", "mist"], flood: ["coolant"], mist: ["coolant"], thru: ["coolant"],
        note: ["comment"], notes: ["comment"], text: ["comment", "manual"],
        search: ["find"], jump: ["find", "go"], filter: ["show", "changed"],
        error: ["messages", "troubleshooting"], problem: ["troubleshooting"], wrong: ["troubleshooting", "check"],
        compare: ["scenarios", "compare"], idea: ["scenarios"], whatif: ["scenarios"],
        doc: ["depth"], stepdown: ["depth", "step"], chip: ["chip", "thinning"],
        install: ["install", "add"], setup: ["install"], column: ["words", "column"], meaning: ["words"]
    };
    var STOP = " a an the to i want how do does my of on in for and or is it its with can what why me be this that at by from as your you make get need please ".split(" ");

    function words(s) {
        return (s || "").toLowerCase().replace(/[^a-z0-9%.:_]+/g, " ").replace(/(^|\s)[.:]+|[.:]+(\s|$)/g, " ")
            .split(" ").filter(function (w) { return w && STOP.indexOf(w) < 0; });
    }

    function esc(s) {
        return s.replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; });
    }

    // How well one entry fits one typed word (with its look-alikes): 0 = not at all.
    function fit(e, w) {
        var best = 0, list = [w].concat(SAME[w] || []);
        list.forEach(function (v, n) {
            var weight = n === 0 ? 1 : 0.6, s = 0;
            var re = new RegExp("(^|[^a-z0-9])" + v.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g");
            if (re.test(e.tl)) s += 12;
            re.lastIndex = 0;
            if (re.test(e.kl)) s += 8;
            re.lastIndex = 0;
            if (re.test(e.hl)) s += 5;
            re.lastIndex = 0;
            var m = e.xl.match(re);
            if (m) s += Math.min(m.length, 6);
            best = Math.max(best, s * weight);
        });
        return best;
    }

    function search(q) {
        var ws = words(q), out = [];
        if (!ws.length || !window.PT_INDEX) return out;
        // Most of the words must be there (people type whole sentences); more words found
        // ranks first, then how well they fit.
        var need = Math.max(1, Math.ceil(ws.length * 0.6));
        window.PT_INDEX.forEach(function (e) {
            if (!e.tl) { e.tl = e.t.toLowerCase(); e.kl = (e.k || "").toLowerCase(); e.hl = (e.h || "").toLowerCase(); e.xl = e.x.toLowerCase(); }
            var total = 0, found = 0;
            ws.forEach(function (w) { var f = fit(e, w); if (f) found++; total += f; });
            if (found >= need && total > 0) out.push({ e: e, s: found * 1000 + total + (e.k ? 2 : 0) });
        });
        out.sort(function (a, b) { return b.s - a.s; });
        // At most two hits per page, twelve in all.
        var per = {}, keep = [];
        out.forEach(function (r) {
            per[r.e.p] = (per[r.e.p] || 0) + 1;
            if (per[r.e.p] <= 2 && keep.length < 12) keep.push(r);
        });
        return keep;
    }

    function snippet(x, ws) {
        var lx = x.toLowerCase(), at = -1;
        ws.forEach(function (w) { var i = lx.indexOf(w); if (i >= 0 && (at < 0 || i < at)) at = i; });
        var from = Math.max(0, at - 60), s = x.substr(from, 190);
        if (from > 0) s = "..." + s.replace(/^\S*\s/, "");
        if (from + 190 < x.length) s = s.replace(/\s\S*$/, "") + "...";
        s = esc(s);
        ws.forEach(function (w) {
            if (w.length < 2) return;
            s = s.replace(new RegExp("(" + esc(w).replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")", "gi"), "<mark>$1</mark>");
        });
        return s;
    }

    function render(list, q) {
        var ws = words(q);
        if (!ws.length) { list.innerHTML = ""; return; }
        var hits = search(q);
        if (!hits.length) {
            list.innerHTML = '<li class="none">Nothing found for &ldquo;' + esc(q) + '&rdquo;. Try fewer words, or a word from the ' +
                '<a href="words.html">word list</a>.</li>';
            return;
        }
        list.innerHTML = hits.map(function (r) {
            var e = r.e, href = e.p + (e.a ? "#" + e.a : "");
            return '<li><a href="' + href + '">' + esc(e.t) + (e.h ? ' <span class="sect">&ndash; ' + esc(e.h) + "</span>" : "") +
                "</a><p>" + snippet(e.x, ws) + "</p></li>";
        }).join("");
    }

    function query() {
        var m = /[?&]q=([^&#]*)/.exec(location.search);
        return m ? decodeURIComponent(m[1].replace(/\+/g, " ")) : "";
    }

    document.addEventListener("DOMContentLoaded", function () {
        // Menu button (phone width).
        var btn = document.querySelector(".menu-btn");
        if (btn) btn.addEventListener("click", function () {
            var open = document.documentElement.classList.toggle("nav-open");
            btn.setAttribute("aria-expanded", open ? "true" : "false");
        });

        // The big search box on the home page: results right under it, as you type.
        var big = document.getElementById("q");
        var bigList = document.getElementById("results");
        if (big && bigList) {
            var go = function () { render(bigList, big.value); };
            big.addEventListener("input", go);
            var q = query();
            if (q) { big.value = q; go(); }
            big.focus();
        }

        // The small box in the top bar: a drop-down of results; Enter opens the home page's search.
        var mini = document.querySelector(".search-mini input");
        var drop = document.querySelector(".search-mini .drop");
        if (mini && drop && !big) {
            var list = drop.querySelector("ul");
            mini.addEventListener("input", function () {
                render(list, mini.value);
                drop.hidden = !mini.value.trim();
            });
            mini.addEventListener("keydown", function (ev) {
                if (ev.key === "Escape") { mini.value = ""; drop.hidden = true; }
            });
            document.addEventListener("click", function (ev) {
                if (!drop.contains(ev.target) && ev.target !== mini) drop.hidden = true;
            });
        }

        // "/" jumps to the search box.
        document.addEventListener("keydown", function (ev) {
            var t = ev.target.tagName;
            if (ev.key === "/" && t !== "INPUT" && t !== "TEXTAREA") {
                var box = big || mini;
                if (box) { ev.preventDefault(); box.focus(); }
            }
        });
    });

    // For tools/make_help.py's check (run in a browser console): search("coolant").
    window.PT_SEARCH = search;
})();
