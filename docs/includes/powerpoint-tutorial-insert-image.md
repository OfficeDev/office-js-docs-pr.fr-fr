<span data-ttu-id="1840b-101">Dans cette étape du didacticiel, vous allez récupérer la photo [Bing](https://www.bing.com) du jour et insérer cette image dans une diapositive.</span><span class="sxs-lookup"><span data-stu-id="1840b-101">In this step of the tutorial, you'll retrieve the [Bing](https://www.bing.com) photo of the day and insert that image into a slide.</span></span>

> [!NOTE]
> <span data-ttu-id="1840b-102">Cette page décrit une étape individuelle du didacticiel sur le complément PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="1840b-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="1840b-103">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément PowerPoint](../tutorials/powerpoint-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="1840b-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="add-the-bing-photo-of-the-day-to-a-slide"></a><span data-ttu-id="1840b-104">Ajout de la photo Bing du jour à une diapositive</span><span class="sxs-lookup"><span data-stu-id="1840b-104">Add the Bing photo of the day to a slide</span></span>

1. <span data-ttu-id="1840b-105">À l’aide de l’explorateur de solutions, ajoutez un nouveau dossier nommé **Controllers** au projet **HelloWorldWeb**.</span><span class="sxs-lookup"><span data-stu-id="1840b-105">Using Solution Explorer, add a new folder named **Controllers** to the **HelloWorldWeb** project.</span></span>

    ![Didacticiel PowerPoint : Fenêtre de l’explorateur de solutions Visual Studio qui met en évidence le dossier Controllers du projet HelloWorldWeb](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. <span data-ttu-id="1840b-107">Cliquez avec le bouton droit de la souris sur le dossier **Controllers**, puis sélectionnez **Ajouter > Nouvel élément généré automatiquement...**.</span><span class="sxs-lookup"><span data-stu-id="1840b-107">Right-click the **Controllers** folder and select **Add > New Scaffolded Item...**.</span></span>

3. <span data-ttu-id="1840b-108">Dans la fenêtre de boîte de dialogue **Ajouter une structure**, sélectionnez **Contrôleur Web API 2 - Vide** et choisissez le bouton **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="1840b-108">In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.</span></span> 

4. <span data-ttu-id="1840b-109">Dans la fenêtre de boîte de dialogue **Ajouter un contrôleur**, saisissez **PhotoController** pour le nom du contrôleur, puis sélectionnez le bouton **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="1840b-109">In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button.</span></span> <span data-ttu-id="1840b-110">Visual Studio crée et ouvre le fichier **PhotoController.cs**.</span><span class="sxs-lookup"><span data-stu-id="1840b-110">Visual Studio creates and opens the **PhotoController.cs** file.</span></span>

5. <span data-ttu-id="1840b-111">Remplacez tout le contenu du fichier **PhotoController.cs** par le code suivant qui appelle le service Bing pour récupérer la photo du jour en tant que chaîne encodée en base 64.</span><span class="sxs-lookup"><span data-stu-id="1840b-111">Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64 encoded string.</span></span> <span data-ttu-id="1840b-112">Lorsque vous utilisez l’API JavaScript Office pour insérer une image dans un document, les données de l’image doivent être spécifiées en tant que chaîne encodée en base 64.</span><span class="sxs-lookup"><span data-stu-id="1840b-112">When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64 encoded string.</span></span>

    ```csharp
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Http;
    using System.Xml;

    namespace HelloWorldWeb.Controllers
    {
        public class PhotoController : ApiController
        {
            public string Get()
            {
                string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

                // Create the request.
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream responseStream = response.GetResponseStream())
                {
                    // Process the result.
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string result = reader.ReadToEnd();

                    // Parse the xml response and to get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64 encoded string.
                    return getPhotoFromURL(photoURL);
                }
            }

            private string getPhotoFromURL(string imageURL)
            {
                var webClient = new WebClient();
                byte[] imageBytes = webClient.DownloadData(imageURL);
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    ```

6. <span data-ttu-id="1840b-113">Dans le fichier **Home.html**, remplacez `TODO1` par le balisage suivant.</span><span class="sxs-lookup"><span data-stu-id="1840b-113">In the **Home.html** file, replace `TODO1` with the following markup.</span></span> <span data-ttu-id="1840b-114">Ce balisage définit le bouton **Insert Image** (Insérer une image) qui s’affichera dans volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="1840b-114">This markup defines the **Insert Image** button that will appear within the add-in's task pane.</span></span>

    ```html
    <button class="ms-Button ms-Button--primary" id="insert-image">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Insert Image</span>
        <span class="ms-Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. <span data-ttu-id="1840b-115">Dans le fichier **Home.js**, remplacez `TODO1` par le code suivant pour attribuer le gestionnaire d’événements pour le bouton **Insert Image** (Insérer une image).</span><span class="sxs-lookup"><span data-stu-id="1840b-115">In the **Home.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

8. <span data-ttu-id="1840b-116">Dans le fichier **Home.js**, remplacez `TODO2` par le code suivant pour définir la fonction **insertImage**.</span><span class="sxs-lookup"><span data-stu-id="1840b-116">In the **Home.js** file, replace `TODO2` with the following code to define the **insertImage** function.</span></span> <span data-ttu-id="1840b-117">Cette fonction extrait l’image du service web Bing, puis appelle la fonction `insertImageFromBase64String` pour insérer cette image dans le document.</span><span class="sxs-lookup"><span data-stu-id="1840b-117">This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.</span></span>

    ```js
    function insertImage() {
        // Get image from from web service (as a Base64 encoded string).
        $.ajax({
            url: "/api/Photo/", success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

9. <span data-ttu-id="1840b-118">Dans le fichier **Home.js**, remplacez `TODO3` par le code suivant pour définir la fonction `insertImageFromBase64String`.</span><span class="sxs-lookup"><span data-stu-id="1840b-118">In the **Home.js** file, replace `TODO3` with the following code to define the `insertImageFromBase64String` function.</span></span> <span data-ttu-id="1840b-119">Cette fonction utilise l’API JavaScript Office pour insérer l’image dans le document.</span><span class="sxs-lookup"><span data-stu-id="1840b-119">This function uses the Office JavaScript API to insert the image into the document.</span></span> <span data-ttu-id="1840b-120">Remarque :</span><span class="sxs-lookup"><span data-stu-id="1840b-120">Note:</span></span> 

    - <span data-ttu-id="1840b-121">l’option `coercionType` spécifiée comme deuxième paramètre de la demande `setSelectedDataAsyc` indique le type de données insérées.</span><span class="sxs-lookup"><span data-stu-id="1840b-121">The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsyc` request indicates the type of data being inserted.</span></span> 

    - <span data-ttu-id="1840b-122">L’objet `asyncResult` encapsule le résultat de la demande `setSelectedDataAsync`, y compris les informations d’état et d’erreur quand la demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="1840b-122">The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.</span></span>

    ```js
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="1840b-123">Tester le complément</span><span class="sxs-lookup"><span data-stu-id="1840b-123">Test the add-in</span></span>

1. <span data-ttu-id="1840b-p107">À l’aide de Visual Studio, testez le nouveau complément PowerPoint en appuyant sur `F5` ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément **Show Taskpane** (Afficher le volet Office) qui apparaît dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="1840b-p107">Using Visual Studio, test the newly created PowerPoint add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Démarrer mis en évidence](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="1840b-127">Dans PowerPoint, sélectionnez le bouton **Show Taskpane** (Afficher le volet Office) dans le ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="1840b-127">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Show Taskpane (Afficher le volet Office) mis en évidence dans le ruban Accueil](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="1840b-129">Dans le volet Office, sélectionnez le bouton **Insert Image** (Insérer une image) permettant d’ajouter la photo Bing du jour sur la diapositive active.</span><span class="sxs-lookup"><span data-stu-id="1840b-129">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.</span></span>

    ![Capture d’écran du complément PowerPoint avec le bouton Insérer une image mis en évidence](../images/powerpoint-tutorial-insert-image-button.png)

4. <span data-ttu-id="1840b-131">Dans Visual Studio, arrêtez le complément en appuyant sur `Shift + F5` ou en choisissant le bouton **Arrêter**.</span><span class="sxs-lookup"><span data-stu-id="1840b-131">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="1840b-132">PowerPoint se ferme automatiquement lorsque le complément est arrêté.</span><span class="sxs-lookup"><span data-stu-id="1840b-132">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)