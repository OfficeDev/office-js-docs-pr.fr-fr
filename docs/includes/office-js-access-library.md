<span data-ttu-id="f707f-101">La bibliothèque de l’interface API JavaScript Office est accessible via le réseau de distribution de contenu (CDN) d’Office JS à l’adresse suivante : `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span><span class="sxs-lookup"><span data-stu-id="f707f-101">The Office JavaScript API library can be accessed via the Office JS content delivery network (CDN) at: `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span></span> <span data-ttu-id="f707f-102">Pour utiliser les API JavaScript Office dans les pages web de votre complément, vous devez référencer le réseau de distribution de contenu dans une balise `<script>`dans la balise `<head>` de la page.</span><span class="sxs-lookup"><span data-stu-id="f707f-102">To use Office JavaScript APIs within any of your add-in's web pages, you must reference the CDN in a `<script>` tag in the `<head>` tag of the page.</span></span>

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> <span data-ttu-id="f707f-103">Pour utiliser les API destinées à la prévisualisation, référencez la version d’évaluation de la bibliothèque de l’interface API JavaScript Office dans le CDN : `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span><span class="sxs-lookup"><span data-stu-id="f707f-103">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

<span data-ttu-id="f707f-104">Si vous souhaitez en savoir plus sur l’accès à la bibliothèque de l’interface API JavaScript pour Office, notamment sur l’obtention d’IntelliSense, consultez [Référencement de la bibliothèque de l’interface API JavaScript pour Office à partir de son réseau de distribution de contenu (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="f707f-104">For more information about accessing the Office JavaScript API library, including how to get IntelliSense, see [Referencing the Office JavaScript API library from its content delivery network (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>