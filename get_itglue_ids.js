function getClientData(clientNames) {
    var clients = [];
    var clientNamesList = JSON.parse(clientNames);
    var rows = document.querySelectorAll("tr[data-id]");
    rows.forEach(function(row) {
        var clientId = row.getAttribute("data-id");
        var clientName = row.querySelector(".column-name") ? row.querySelector(".column-name").textContent : "N/A";
        if (clientNamesList.includes(clientName)) {
            clients.push({id: clientId, name: clientName});
        }
    });
    return clients;
}