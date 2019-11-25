
export let config = {

    version: "0.0.1",
    url: "https://localhost:3000",
    vorlon: "https://localhost:1337",

    id: function () {
        return '{xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx}'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

};