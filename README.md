# Html2Word
C# Web API to get a complete Word document from a HTML string.

## Example how to call the API in Angular
```
$http.post(baseUrl + "/api/document/generate", { Html : $scope.document.dataUri })
    .then(function success(response) {
        window.open(response.data);
    }, function error(response) {
        console.log("ERROR! " + response.statusText);
    });
```
