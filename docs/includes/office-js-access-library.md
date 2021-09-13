La bibliothèque de l’interface API JavaScript Office est accessible via le réseau de distribution de contenu (CDN) d’Office JS à l’adresse suivante : `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`. Pour utiliser les API JavaScript Office dans les pages web de votre complément, vous devez référencer le réseau de distribution de contenu dans une balise `<script>`dans la balise `<head>` de la page.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> Pour utiliser les API destinées à la prévisualisation, référencez la version d’évaluation de la bibliothèque de l’interface API JavaScript Office dans le CDN : `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

Si vous souhaitez en savoir plus sur l’accès à la bibliothèque de l’interface API JavaScript pour Office, notamment sur l’obtention d’IntelliSense, consultez [Référencement de la bibliothèque de l’interface API JavaScript pour Office à partir de son réseau de distribution de contenu (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).