# <a name="get-the-whole-document-from-an-add-in-for-powerpoint-or-word"></a><span data-ttu-id="087db-101">Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word</span><span class="sxs-lookup"><span data-stu-id="087db-101">Get the whole document from an add-in for PowerPoint or Word</span></span>

<span data-ttu-id="087db-p101">Vous pouvez créer un Complément Office pour permettre l’envoi ou la publication en un clic d’un document Word 2013 ou PowerPoint 2013 sur un emplacement distant. Cet article explique comment créer un complément du volet de tâches pour PowerPoint 2013 qui envoie les données d’une présentation sous la forme d’un objet de données à un serveur web via une requête HTTP.</span><span class="sxs-lookup"><span data-stu-id="087db-p101">You can create an Office Add-in to provide one-click sending or publishing of a Word 2013 or PowerPoint 2013 document to a remote location. This article demonstrates how to build a simple task pane add-in for PowerPoint 2013 that gets all of the presentation as a data object and sends that data to a web server via an HTTP request.</span></span>

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a><span data-ttu-id="087db-104">Éléments requis pour créer un complément pour PowerPoint ou Word</span><span class="sxs-lookup"><span data-stu-id="087db-104">Prerequisites for creating an add-in for PowerPoint or Word</span></span>

<span data-ttu-id="087db-p102">Dans cet article, vous utilisez un éditeur de texte pour créer le complément du volet Office pour PowerPoint ou Word. Pour créer le complément du volet Office, vous devez créer les fichiers suivants :</span><span class="sxs-lookup"><span data-stu-id="087db-p102">This article assumes that you are using a text editor to create the task pane add-in for PowerPoint or Word. To create the task pane add-in, you must create the following files:</span></span>

- <span data-ttu-id="087db-107">Sur un dossier réseau partagé ou sur un serveur web, vous avez besoin des fichiers suivants :</span><span class="sxs-lookup"><span data-stu-id="087db-107">On a shared network folder or on a web server, you need the following files:</span></span>
    
    - <span data-ttu-id="087db-108">Un fichier HTML (GetDoc_App.html) qui contient l’interface utilisateur, ainsi que les liens vers les fichiers JavaScript (notamment office.js et fichiers .js propres à l’hôte) et les fichiers CSS (Cascading Style Sheet).</span><span class="sxs-lookup"><span data-stu-id="087db-108">An HTML file (GetDoc_App.html) that contains the user interface plus links to the JavaScript files (including office.js and host-specific .js files) and Cascading Style Sheet (CSS) files.</span></span>
           
    - <span data-ttu-id="087db-109">Un fichier JavaScript (GetDoc_App.js) qui contient la logique de programmation du complément.</span><span class="sxs-lookup"><span data-stu-id="087db-109">A JavaScript file (GetDoc_App.js) to contain the programming logic of the add-in.</span></span>
    
    - <span data-ttu-id="087db-110">Un fichier CSS (Program.css) qui contient les styles et la mise en forme du complément.</span><span class="sxs-lookup"><span data-stu-id="087db-110">A CSS file (Program.css) to contain the styles and formatting for the add-in.</span></span>
    
- <span data-ttu-id="087db-p103">Un fichier manifeste XML (GetDoc_App.xml) pour le complément, disponible dans un dossier réseau partagé ou un catalogue de compléments. Le fichier manifeste doit pointer vers l’emplacement du fichier HTML mentionné précédemment.</span><span class="sxs-lookup"><span data-stu-id="087db-p103">An XML manifest file (GetDoc_App.xml) for the add-in, available on a shared network folder or add-in catalog. The manifest file must point to the location of the HTML file mentioned previously.</span></span>
    
<span data-ttu-id="087db-113">Vous pouvez également créer un complément pour PowerPoint à l’aide de [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visual-studio) ou de [n’importe quel éditeur](../quickstarts/powerpoint-quickstart.md?tabs=visual-studio-code), ou pour Word à l’aide de [Visual Studio](../quickstarts/word-quickstart.md?tabs=visual-studio) ou de [n’importe quel éditeur](../quickstarts/word-quickstart.md?tabs=visual-studio-code).</span><span class="sxs-lookup"><span data-stu-id="087db-113">You can also create an add-in for PowerPoint by using [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visual-studio) or [any editor](../quickstarts/powerpoint-quickstart.md?tabs=visual-studio-code) or for Word by using [Visual Studio](../quickstarts/word-quickstart.md?tabs=visual-studio) or [any editor](../quickstarts/word-quickstart.md?tabs=visual-studio-code).</span></span> 

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a><span data-ttu-id="087db-114">Concepts fondamentaux à connaître pour créer un complément du volet Office</span><span class="sxs-lookup"><span data-stu-id="087db-114">Core concepts to know for creating a task pane add-in</span></span>

<span data-ttu-id="087db-p104">Avant de commencer à créer ce complément pour PowerPoint ou Word, vous devez savoir comment créer des Compléments Office et utiliser des requêtes HTTP. Cet article n’explique pas comment décoder du texte codé en Base64 à partir d’une requête HTTP sur un serveur web.</span><span class="sxs-lookup"><span data-stu-id="087db-p104">Before you begin creating this add-in for PowerPoint or Word, you should be familiar with building Office Add-ins and working with HTTP requests. This article does not discuss how to decode Base64-encoded text from an HTTP request on a web server.</span></span> 

## <a name="create-the-manifest-for-the-add-in"></a><span data-ttu-id="087db-117">Créer le manifeste pour le complément</span><span class="sxs-lookup"><span data-stu-id="087db-117">Create the manifest for the add-in</span></span>


<span data-ttu-id="087db-118">Le fichier manifeste XML pour le complément PowerPoint fournit des informations importantes sur le complément : les applications qui peuvent l’héberger, l’emplacement du fichier HTML, le titre et la description du complément, et bien d’autres caractéristiques.</span><span class="sxs-lookup"><span data-stu-id="087db-118">The XML manifest file for the add-in for PowerPoint provides important information about the add-in: what applications can host it, the location of the HTML file, the add-in title and description, and many other characteristics.</span></span>

1. <span data-ttu-id="087db-119">Dans l’éditeur de texte, ajoutez le code suivant au fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="087db-119">In a text editor, add the following code to the manifest file.</span></span>
    
    ```xml  
    <?xml version="1.0" encoding="utf-8" ?> 
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
    xsi:type="TaskPaneApp">
        <Id>[Replace_With_Your_GUID]</Id> 
        <Version>1.0</Version> 
        <ProviderName>[Provider Name]</ProviderName> 
        <DefaultLocale>EN-US</DefaultLocale> 
        <DisplayName DefaultValue="Get Doc add-in" /> 
        <Description DefaultValue="My get PowerPoint or Word document add-in." /> 
        <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" /> 
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

2. <span data-ttu-id="087db-120">Enregistrez le fichier avec le nom GetDoc_App.xml et un encodage UTF-8 sur un emplacement réseau ou dans un catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="087db-120">Save the file as GetDoc_App.xml using UTF-8 encoding to a network location or to an add-in catalog.</span></span>
    
## <a name="create-the-user-interface-for-the-add-in"></a><span data-ttu-id="087db-121">Créer l’interface utilisateur pour le complément</span><span class="sxs-lookup"><span data-stu-id="087db-121">Create the user interface for the add-in</span></span>

<span data-ttu-id="087db-p105">Pour l’interface utilisateur du complément, vous pouvez utiliser du code HTML, écrit directement dans le fichier GetDoc_App.html. La logique de programmation et les fonctionnalités du complément doivent être contenues dans un fichier JavaScript (par exemple, GetDoc_App.js).</span><span class="sxs-lookup"><span data-stu-id="087db-p105">For the user interface of the add-in, you can use HTML, written directly into the GetDoc_App.html file. The programming logic and functionality of the add-in must be contained in a JavaScript file (for example, GetDoc_App.js).</span></span>

<span data-ttu-id="087db-124">Utilisez la procédure suivante pour créer une interface utilisateur simple pour le complément, avec un titre et un bouton unique.</span><span class="sxs-lookup"><span data-stu-id="087db-124">Use the following procedure to create a simple user interface for the add-in that includes a heading and a single button.</span></span>

1. <span data-ttu-id="087db-125">Dans un nouveau fichier dans l’éditeur de texte, ajoutez le code HTML suivant.</span><span class="sxs-lookup"><span data-stu-id="087db-125">In a new file in the text editor, add the following HTML.</span></span>
        
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

2. <span data-ttu-id="087db-126">Enregistrez le fichier avec le nom GetDoc_App.html et un encodage UTF-8 sur un emplacement réseau ou sur un serveur web.</span><span class="sxs-lookup"><span data-stu-id="087db-126">Save the file as GetDoc_App.html using UTF-8 encoding to a network location or to a web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="087db-127">Assurez-vous que les balises **head** du complément contiennent une balise **script** et un lien valide vers le fichier office.js.</span><span class="sxs-lookup"><span data-stu-id="087db-127">Be sure that the **head** tags of the add-in contains a **script** tag with a valid link to the office.js file.</span></span> 

    <span data-ttu-id="087db-p106">Nous allons utiliser des styles CSS pour donner au complément une apparence simple, moderne et professionnelle. Utilisez le code CSS suivant pour définir le style du complément.</span><span class="sxs-lookup"><span data-stu-id="087db-p106">We'll use some CSS to give the add-in a simple, yet modern and professional appearance. Use the following CSS to define the style of the add-in.</span></span>

3. <span data-ttu-id="087db-130">Dans un nouveau fichier dans l’éditeur de texte, ajoutez le code CSS suivant.</span><span class="sxs-lookup"><span data-stu-id="087db-130">In a new file in the text editor, add the following CSS.</span></span>
        
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

4. <span data-ttu-id="087db-131">Enregistrez le fichier avec le nom Program.css et un encodage UTF-8 sur l’emplacement réseau ou sur le serveur web sur lequel se trouve le fichier GetDoc_App.html.</span><span class="sxs-lookup"><span data-stu-id="087db-131">Save the file as Program.css using UTF-8 encoding to the network location or to the web server where the GetDoc_App.html file is located.</span></span>
    
## <a name="add-the-javascript-to-get-the-document"></a><span data-ttu-id="087db-132">Ajouter le code JavaScript pour obtenir le document</span><span class="sxs-lookup"><span data-stu-id="087db-132">Add the JavaScript to get the document</span></span>

<span data-ttu-id="087db-133">Dans le code pour le complément, un gestionnaire vers l’événement [Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) ajoute un gestionnaire à l’événement Click du bouton **Envoyer** du formulaire et informe l’utilisateur que le complément est prêt.</span><span class="sxs-lookup"><span data-stu-id="087db-133">In the code for the add-in, a handler to the [Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) event adds a handler to the click event of the **Submit** button on the form and informs the user that the add-in is ready.</span></span>

<span data-ttu-id="087db-134">L’exemple de code suivant montre le gestionnaire d’événements pour l’événement  **Office.initialize** avec une fonction d’aide, `updateStatus`, pour écrire dans status div.</span><span class="sxs-lookup"><span data-stu-id="087db-134">The following code example shows the event handler for the  **Office.initialize** event along with a helper function, `updateStatus`, for writing to the status div.</span></span>

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
    statusInfo.innerHTML += message + "<br/>";
}
```

<span data-ttu-id="087db-p107">Quand vous cliquez sur le bouton  **Envoyer** dans l’interface utilisateur, le complément appelle la fonction `sendFile`, qui contient un appel à la méthode [Document.getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-). La méthode  **getFileAsync** utilise le modèle asynchrone, comme d’autres méthodes de l’API JavaScript pour Office. Elle utilise un paramètre obligatoire, _fileType_, et deux paramètres facultatifs,  _options_ et _callback_.</span><span class="sxs-lookup"><span data-stu-id="087db-p107">When you choose the  **Submit** button in the UI, the add-in calls the `sendFile` function, which contains a call to the [Document.getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-) method. The **getFileAsync** method uses the asynchronous pattern, similar to other methods in the JavaScript API for Office. It has one required parameter, _fileType_, and two optional parameters,  _options_ and _callback_.</span></span> 

<span data-ttu-id="087db-p108">Le paramètre _fileType_ attend l’une des trois constantes de l’énumération [FileType](https://docs.microsoft.com/javascript/api/office/office.filetype?view=office-js) :  **Office.FileType.Compressed** (« compressé »), **Office.FileType.PDF** (« PDF ») ou **Office.FileType.Text** (« texte »). PowerPoint prend en charge uniquement **Compressed** comme argument, tandis que Word prend en charge les trois. Lorsque vous transmettez **Compressed** pour le paramètre _fileType_, la méthode  **getFileAsync** renvoie le document sous la forme d’un fichier de présentation PowerPoint 2013 (*.pptx) ou Word 2013 d’un fichier de document (*.docx) en créant une copie temporaire du fichier sur l’ordinateur local.</span><span class="sxs-lookup"><span data-stu-id="087db-p108">The  _fileType_ parameter expects one of three constants from the [FileType](https://docs.microsoft.com/javascript/api/office/office.filetype?view=office-js) enumeration: **Office.FileType.Compressed** ("compressed"), **Office.FileType.PDF** ("pdf"), or **Office.FileType.Text** ("text"). PowerPoint supports only **Compressed** as an argument; Word supports all three. When you pass in **Compressed** for the _fileType_ parameter, the **getFileAsync** method returns the document as a PowerPoint 2013 presentation file (*.pptx) or Word 2013 document file (*.docx) by creating a temporary copy of the file on the local computer.</span></span>

<span data-ttu-id="087db-p109">La méthode **getFileAsync** renvoie une référence au fichier sous la forme d’un objet [File](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js). L’objet **File** expose quatre membres : la propriété [size](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#size), la propriété [sliceCount](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#slicecount), la méthode [getSliceAsync](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#getsliceasync-sliceindex--callback-) et la méthode [closeAsync](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#closeasync-callback-). La propriété **size** renvoie le nombre d’octets du fichier. La propriété **sliceCount** renvoie le nombre d’objets [Slice](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js) (expliqué plus loin dans cet article) du fichier.</span><span class="sxs-lookup"><span data-stu-id="087db-p109">The  **getFileAsync** method returns a reference to the file as a [File](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js) object. The **File** object exposes four members: the [size](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#size) property, [sliceCount](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#slicecount) property, [getSliceAsync](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#getsliceasync-sliceindex--callback-) method, and [closeAsync](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#closeasync-callback-) method. The **size** property returns the number of bytes in the file. The **sliceCount** returns the number of [Slice](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js) objects (discussed later in this article) in the file.</span></span>

<span data-ttu-id="087db-p110">Utilisez le code suivant pour obtenir le document Word ou PowerPoint sous la forme d’un objet  **File** en utilisant la méthode **Document.getFileAsync**, puis appeler la fonction  `getSlice` définie localement. L’objet **File**, une variable compteur et le nombre total de tranches du fichier sont passés avec l’appel à  `getSlice` dans un objet anonyme.</span><span class="sxs-lookup"><span data-stu-id="087db-p110">Use the following code to get the PowerPoint or Word document as a  **File** object using the **Document.getFileAsync** method and then makes a call to the locally defined `getSlice` function. Note that the **File** object, a counter variable, and the total number of slices in the file are passed along in the call to `getSlice` in an anonymous object.</span></span>

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

<span data-ttu-id="087db-p111">La fonction locale  `getSlice` appelle la méthode **File.getSliceAsync** pour extraire une tranche de l’objet **File**. La méthode  **getSliceAsync** retourne un objet **Slice** de la collection de tranches. Elle a deux paramètres requis, _sliceIndex_ et _callback_. Le paramètre  _sliceIndex_ utilise un entier comme indexeur dans la collection de tranches. Comme d’autres fonctions de l’API JavaScript pour Office, la méthode **getSliceAsync** prend également une fonction de rappel comme paramètre pour gérer les résultats de l’appel de la méthode.</span><span class="sxs-lookup"><span data-stu-id="087db-p111">The local function  `getSlice` makes a call to the **File.getSliceAsync** method to retrieve a slice from the **File** object. The **getSliceAsync** method returns a **Slice** object from the collection of slices. It has two required parameters, _sliceIndex_ and _callback_. The  _sliceIndex_ parameter takes an integer as an indexer into the collection of slices. Like other functions in the JavaScript API for Office, the **getSliceAsync** method also takes a callback function as a parameter to handle the results from the method call.</span></span>

<span data-ttu-id="087db-152">L’objet **Section** vous donne accès aux données contenues dans le fichier.</span><span class="sxs-lookup"><span data-stu-id="087db-152">The **Slice** object gives you access to the data contained in the file.</span></span> <span data-ttu-id="087db-153">Sauf indication contraire dans le paramètre _options_ de la méthode **getFileAsync** , l’objet **Section** a une taille de 4 Mo.</span><span class="sxs-lookup"><span data-stu-id="087db-153">Unless otherwise specified in the _options_ parameter of the **getFileAsync** method, the **Slice** object is 4 MB in size.</span></span> <span data-ttu-id="087db-154">L’objet **Section**  expose trois propriétés : [taille](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js#size), [données](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js#data), et [index](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js#index).</span><span class="sxs-lookup"><span data-stu-id="087db-154">The **Slice** object exposes three properties: [size](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js#size), [data](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js#data), and [index](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js#index).</span></span> <span data-ttu-id="087db-155">La propriété **taille** obtient la taille, en octets, de la section.</span><span class="sxs-lookup"><span data-stu-id="087db-155">The **size** property gets the size, in bytes, of the slice.</span></span> <span data-ttu-id="087db-156">La propriété **index** obtient un entier qui représente la position de la section dans la collection de sections.</span><span class="sxs-lookup"><span data-stu-id="087db-156">The **index** property gets an integer that represents the slice's position in the collection of slices.</span></span>

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

<span data-ttu-id="087db-p113">La propriété  **Slice.data** retourne les données brutes du fichier sous la forme d’un tableau d’octets. Si les données sont au format texte (c’est-à-dire, XML ou texte brut), la tranche contient du texte brut. Si vous avez passé **Office.FileType.Compressed** pour le paramètre _fileType_ de **Document.getFileAsync**, la tranche contient les données binaires du fichier sous la forme d’un tableau d’octets. Dans le cas d’un fichier PowerPoint ou Word, les tranches contiennent des tableaux d’octets.</span><span class="sxs-lookup"><span data-stu-id="087db-p113">The  **Slice.data** property returns the raw data of the file as a byte array. If the data is in text format (that is, XML or plain text), the slice contains the raw text. If you pass in **Office.FileType.Compressed** for the _fileType_ parameter of **Document.getFileAsync**, the slice contains the binary data of the file as a byte array. In the case of a PowerPoint or Word file, the slices contain byte arrays.</span></span>

<span data-ttu-id="087db-p114">Vous devez implémenter votre propre fonction (ou utiliser une bibliothèque disponible) pour convertir les données d’un tableau d’octets en chaîne codée en Base64. Pour plus d’informations sur le codage en Base64 avec JavaScript, voir [Codage et décodage en Base64](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).</span><span class="sxs-lookup"><span data-stu-id="087db-p114">You must implement your own function (or use an available library) to convert byte array data to a Base64-encoded string. For information about Base64 encoding with JavaScript, see [Base64 encoding and decoding](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).</span></span>

<span data-ttu-id="087db-163">Une fois les données converties en Base64, vous pouvez les transmettre au serveur web de plusieurs façons, notamment dans le corps d’une demande POST HTTP.</span><span class="sxs-lookup"><span data-stu-id="087db-163">Once you have converted the data to Base64, you can then transmit it to a web server in several ways -- including as the body of an HTTP POST request.</span></span>

<span data-ttu-id="087db-164">Ajoutez le code suivant pour envoyer une tranche au service web.</span><span class="sxs-lookup"><span data-stu-id="087db-164">Add the following code to send a slice to a web service.</span></span>

> [!NOTE]
> <span data-ttu-id="087db-p115">Ce code envoie un fichier PowerPoint ou Word sur le serveur web en plusieurs tranches. Le serveur web ou le service doit compiler chaque tranche en un unique fichier .pptx avant de pouvoir l’utiliser.</span><span class="sxs-lookup"><span data-stu-id="087db-p115">This code sends a PowerPoint or Word file to the web server in multiple slices. The web server or service must compile each individual slice into a single .pptx file before you can perform any manipulations on it.</span></span>

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

<span data-ttu-id="087db-p116">Comme son nom l’indique, la méthode  **File.closeAsync** ferme la connexion au document et libère les ressources. Bien que le garbage sandbox des Compléments Office collecte les références hors étendue aux fichiers, il est conseillé de fermer explicitement les fichiers quand le code a terminé de les utiliser. La méthode **closeAsync** a un unique paramètre, _callback_, qui spécifie la fonction à appeler à la fin de l’appel.</span><span class="sxs-lookup"><span data-stu-id="087db-p116">As the name implies, the  **File.closeAsync** method closes the connection to the document and frees up resources. Although the Office Add-ins sandbox garbage collects out-of-scope references to files, it is still a best practice to explicitly close files once your code is done with them. The **closeAsync** method has a single parameter, _callback_, that specifies the function to call on the completion of the call.</span></span>

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