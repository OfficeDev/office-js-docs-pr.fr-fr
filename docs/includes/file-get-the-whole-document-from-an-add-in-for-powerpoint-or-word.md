Vous pouvez créer un Complément Office pour permettre l’envoi ou la publication en un clic d’un document Word 2013 ou PowerPoint 2013 sur un emplacement distant. Cet article explique comment créer un complément du volet de tâches pour PowerPoint 2013 qui envoie les données d’une présentation sous la forme d’un objet de données à un serveur web via une requête HTTP.

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a>Éléments requis pour créer un complément pour PowerPoint ou Word

Dans cet article, vous utilisez un éditeur de texte pour créer le complément du volet Office pour PowerPoint ou Word. Pour créer le add-in du volet Des tâches, vous devez créer les fichiers suivants.

- Sur un dossier réseau partagé ou sur un serveur web, vous avez besoin des fichiers suivants.

  - Fichier HTML (GetDoc_App.html) qui contient l’interface utilisateur, ainsi que des liens vers les fichiers JavaScript (y compris les fichiers office.js et .js propres à l’application) et les fichiers CSS (Cascading Style Sheet).

  - Un fichier JavaScript (GetDoc_App.js) qui contient la logique de programmation du complément.

  - Un fichier CSS (Program.css) qui contient les styles et la mise en forme du complément.

- Un fichier manifeste XML (GetDoc_App.xml) pour le complément, disponible dans un dossier réseau partagé ou un catalogue de compléments. Le fichier manifeste doit pointer vers l’emplacement du fichier HTML mentionné précédemment.

Vous pouvez également créer un module pour PowerPoint à l’aide de [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio) ou du générateur [Yeoman](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) pour les Office ou pour Word à l’aide du générateur [Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio) ou [Yeoman](../quickstarts/word-quickstart.md?tabs=yeomangenerator)pour les Office.

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a>Concepts fondamentaux à connaître pour créer un complément du volet Office

Avant de commencer à créer ce complément pour PowerPoint ou Word, vous devez savoir comment créer des Compléments Office et utiliser des requêtes HTTP. Cet article n’explique pas comment décoder du texte codé en Base64 à partir d’une requête HTTP sur un serveur web.

## <a name="create-the-manifest-for-the-add-in"></a>Créer le manifeste pour le complément

Le fichier manifeste XML pour le complément PowerPoint fournit des informations importantes sur le complément : les applications qui peuvent l’héberger, l’emplacement du fichier HTML, le titre et la description du complément, et bien d’autres caractéristiques.

1. Dans l’éditeur de texte, ajoutez le code suivant au fichier manifeste.

    ```xml  
    <?xml version="1.0" encoding="utf-8" ?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xsi:type="TaskPaneApp">
        <Id>[Replace_With_Your_GUID]</Id>
        <Version>1.0</Version>
        <ProviderName>[Provider Name]</ProviderName>
        <DefaultLocale>EN-US</DefaultLocale>
        <DisplayName DefaultValue="Get Doc add-in" />
        <Description DefaultValue="My get PowerPoint or Word document add-in." />
        <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
        <Hosts>
        <Host Name="Document" />
        <Host Name="Presentation" />
        </Hosts>
        <DefaultSettings>
        <SourceLocation DefaultValue="[Network location of app]/GetDoc_App.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

2. Enregistrez le fichier avec le nom GetDoc_App.xml et un encodage UTF-8 sur un emplacement réseau ou dans un catalogue de compléments.

## <a name="create-the-user-interface-for-the-add-in"></a>Créer l’interface utilisateur pour le complément

Pour l’interface utilisateur du complément, vous pouvez utiliser du code HTML, écrit directement dans le fichier GetDoc_App.html. La logique de programmation et les fonctionnalités du complément doivent être contenues dans un fichier JavaScript (par exemple, GetDoc_App.js).

Utilisez la procédure suivante pour créer une interface utilisateur simple pour le complément, avec un titre et un bouton unique.

1. Dans un nouveau fichier dans l’éditeur de texte, ajoutez le code HTML suivant.

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
            <title>Publish presentation</title>
            <link rel="stylesheet" type="text/css" href="Program.css" />
            <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js" type="text/javascript"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="GetDoc_App.js"></script>
        </head>
        <body>
        <form>
            <h1>Publish presentation</h1>
            <br />
            <div><input id='submit' type="button" value="Submit" /></div>
            <br />
            <div><h2>Status</h2> 
                <div id="status"></div>
            </div>
        </form>
        </body>
    </html>
    ```

2. Enregistrez le fichier avec le nom GetDoc_App.html et un encodage UTF-8 sur un emplacement réseau ou sur un serveur web.

    > [!NOTE]
    > Assurez-vous que les balises **head** du complément contiennent une balise **script** et un lien valide vers le fichier office.js.

    Nous allons utiliser des styles CSS pour donner au complément une apparence simple, moderne et professionnelle. Utilisez le code CSS suivant pour définir le style du complément.

3. Dans un nouveau fichier dans l’éditeur de texte, ajoutez le code CSS suivant.

    ```css  
    body
    {
        font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
    }
    h1,h2
    {
        text-decoration-color:#4ec724;
    }
    input [type="submit"], input[type="button"]
    {
        height:24px;
        padding-left:1em;
        padding-right:1em;
        background-color:white;
        border:1px solid grey;
        border-color: #dedfe0 #b9b9b9 #b9b9b9 #dedfe0;
        cursor:pointer;
    }
    ```

4. Enregistrez le fichier avec le nom Program.css et un encodage UTF-8 sur l’emplacement réseau ou sur le serveur web sur lequel se trouve le fichier GetDoc_App.html.

## <a name="add-the-javascript-to-get-the-document"></a>Ajouter le code JavaScript pour obtenir le document

Dans le code pour le complément, un gestionnaire vers l’événement [Office.initialize](/javascript/api/office) ajoute un gestionnaire à l’événement Click du bouton **Envoyer** du formulaire et informe l’utilisateur que le complément est prêt.

L’exemple de code suivant montre le handler d’événement pour l’événement, ainsi qu’une fonction d’aide, pour l’écriture dans la `Office.initialize` `updateStatus` div d’état.

```js
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

        // Execute sendFile when submit is clicked
        $('#submit').click(function () {
            sendFile();
        });

        // Update status
        updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div.
function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo[0].innerHTML += message + "<br/>";
}
```

Lorsque vous choisissez **le** bouton Envoyer dans l’interface utilisateur, le add-in appelle la fonction, qui contient un appel à la méthode `sendFile` [Document.getFileAsync.](/javascript/api/office/office.document#getFileAsync_fileType__options__callback_) La méthode utilise le modèle asynchrone, semblable à d’autres méthodes dans `getFileAsync` l’API JavaScript pour Office. Elle utilise un paramètre obligatoire, _fileType_, et deux paramètres facultatifs,  _options_ et _callback_.

Le _paramètre fileType_ attend l’une des trois constantes de l’éumération [FileType](/javascript/api/office/office.filetype) : (« compressé »)Office.FileType.PDF(« pdf ») ou `Office.FileType.Compressed`  **Office. FileType.Text** (« text »). La prise en charge actuelle des types de fichiers pour chaque plateforme est répertoriée sous les remarques [Document.getFileType.](/javascript/api/office/office.document#getFileAsync_fileType__callback_) Lorsque vous passez  compressé pour le paramètre _fileType,_ la méthode renvoie le document en tant que fichier de présentation `getFileAsync` PowerPoint 2013 *(.pptx) ou Word 2013 (.docx)* en créant une copie temporaire du fichier sur l’ordinateur local.

La `getFileAsync` méthode renvoie une référence au fichier en tant [qu’objet](/javascript/api/office/office.file) File. L’objet expose quatre membres : la propriété `File` [size,](/javascript/api/office/office.file#size) la propriété [sliceCount,](/javascript/api/office/office.file#sliceCount) la méthode [getSliceAsync](/javascript/api/office/office.file#getSliceAsync_sliceIndex__callback_) et [la méthode closeAsync.](/javascript/api/office/office.file#closeAsync_callback_) La `size` propriété renvoie le nombre d’octets dans le fichier. Renvoie `sliceCount` le nombre [d’objets Slice](/javascript/api/office/office.slice) (décrits plus loin dans cet article) dans le fichier.

Utilisez le code suivant pour obtenir le document PowerPoint ou Word en tant qu’objet à l’aide de la méthode, puis effectuez un appel à la `File` `Document.getFileAsync` fonction définie `getSlice` localement. Notez que l’objet, une variable de compteur et le nombre total de tranches du fichier sont transmis dans l’appel à un `File` `getSlice` objet anonyme.

```js
// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {
    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status == Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size + " bytes");
                getSlice(state);
            }
            else {
                updateStatus(result.status);
            }
        });
}
```

La fonction locale `getSlice` appelle la méthode pour récupérer une section de `File.getSliceAsync` `File` l’objet. La `getSliceAsync` méthode renvoie un objet de la collection de `Slice` tranches. Elle a deux paramètres requis, _sliceIndex_ et _callback_. Le paramètre  _sliceIndex_ utilise un entier comme indexeur dans la collection de tranches. Comme d’autres fonctions dans l’INTERFACE API JavaScript pour Office, la méthode prend également une fonction de rappel comme paramètre pour gérer les résultats de l’appel `getSliceAsync` de méthode.
ion appelle la méthode `getSlice` **File.getSliceAsync** pour récupérer une tranche de **l’objet** File. La méthode  **getSliceAsync** retourne un objet **Slice** de la collection de tranches. Elle a deux paramètres requis, _sliceIndex_ et _callback_. Le paramètre  _sliceIndex_ utilise un entier comme indexeur dans la collection de tranches. Comme les autres fonctions de l’API JavaScript Office, la méthode **getSliceAsync** prend également une fonction de rappel comme paramètre pour gérer les résultats de l’appel de méthode.

`Slice`L’objet vous donne accès aux données contenues dans le fichier. Sauf indication contraire dans le paramètre _options_ de la méthode, la taille de l’objet `getFileAsync` est de `Slice` 4 Mo. `Slice`L’objet expose trois propriétés : [taille,](/javascript/api/office/office.slice#size) [données](/javascript/api/office/office.slice#data)et [index](/javascript/api/office/office.slice#index). La `size` propriété obtient la taille, en octets, de la tranche. La propriété obtient un nombre integer qui représente la position de la tranche `index` dans la collection de tranches.

```js
// Get a slice from the file and then call sendSlice.
function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}
```

La `Slice.data` propriété renvoie les données brutes du fichier sous la mesure d’un tableau d’byte. Si les données sont au format texte (c’est-à-dire, XML ou texte brut), la tranche contient du texte brut. Si vous passez **Office.FileType.Compressed** pour le paramètre _fileType_ de , la tranche contient les données binaires du fichier sous forme de tableau d’byte. `Document.getFileAsync` Dans le cas d’un fichier PowerPoint ou Word, les tranches contiennent des tableaux d’octets.

Vous devez implémenter votre propre fonction (ou utiliser une bibliothèque disponible) pour convertir les données d’un tableau d’octets en chaîne codée en Base64. Pour plus d’informations sur le codage en Base64 avec JavaScript, voir [Codage et décodage en Base64](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).

Une fois les données converties en Base64, vous pouvez les transmettre au serveur web de plusieurs façons, notamment dans le corps d’une demande POST HTTP.

Ajoutez le code suivant pour envoyer une tranche au service web.

> [!NOTE]
> Ce code envoie un fichier PowerPoint word au serveur web en plusieurs tranches. Le serveur web ou le service doit appender chaque tranche individuelle dans un fichier unique, puis l’enregistrer en tant que fichier .pptx ou .docx, avant de pouvoir y effectuer des manipulations.

```js
function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
        var fileData = myEncodeBase64(data);

        // Create a new HTTP request. You need to send the request
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();

        // Create a handler function to update the status
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                }
                else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "[Your receiving page or service]");
        request.setRequestHeader("Slice-Number", slice.index);

        // Send the file as the body of an HTTP POST
        // request to the web server.
        request.send(fileData);
    }
}
```

Comme son nom l’indique, la méthode ferme la connexion au `File.closeAsync` document et libère des ressources. Bien que le garbage sandbox des Compléments Office collecte les références hors étendue aux fichiers, il est conseillé de fermer explicitement les fichiers quand le code a terminé de les utiliser. La `closeAsync` méthode a un seul paramètre, _callback,_ qui spécifie la fonction à appeler à la fin de l’appel.

```js
function closeFile(state) {
    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status == "succeeded") {
            updateStatus("File closed.");
        }
        else {
            updateStatus("File couldn't be closed.");
        }
    });
}
```