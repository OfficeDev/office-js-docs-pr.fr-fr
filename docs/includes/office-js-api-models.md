L’API JavaScript pour Office inclut deux modèles distincts :

- **Les API spécifiques aux hôtes** fournissent des objets fortement typés qui peuvent être utilisés pour interagir avec des objets natifs d’une application Office spécifique. Par exemple, vous pouvez utiliser les API JavaScript pour Excel pour accéder à des feuilles de calcul, plages, tableaux, graphiques, etc. Les API spécifiques aux hôtes sont actuellement disponibles pour les hôtes suivants :

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)

    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)

    Ce modèle API utilise [promet](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) et vous permet de spécifier plusieurs opérations dans chaque demande que vous envoyez à l’hôte Office. Les opérations de traitement par lots de cette manière peuvent améliorer sensiblement les performances des compléments dans les applications Office sur le Web. Les API spécifiques aux hôtes ont été introduites avec Office 2016 et ne peuvent pas être utilisées pour interagir avec Office 2013.

    > [!NOTE]
    > Il existe également une API spécifique de l’hôte pour [Visio](../reference/overview/visio-javascript-reference-overview.md), mais vous pouvez l’utiliser uniquement dans les pages SharePoint Online pour interagir avec les diagrammes Visio incorporés dans la page. Les compléments Web Office ne sont pas pris en charge dans Visio.

- Les API **Communes** peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office. Ce modèle API utilise des[rappels](https://developer.mozilla.org/docs/Glossary/Callback_function), qui vous permettent de spécifier une seule opération dans chaque demande envoyée à l’hôte Office. Les API communes ont été introduites avec Office 2013 et peuvent être utilisées pour interagir avec Office 2013 ou version ultérieure. Si vous souhaitez plus en savoir sur le modèle objet API commun, qui inclut des API pour l’interaction avec Outlook et PowerPoint, veuillez consulter [Modèle d’objet API JavaScript communes](../develop/office-javascript-api-object-model.md).

> [!NOTE]
> Certaines fonctions Excel personnalisées s’exécutent dans le cadre d’une exécution unique qui hiérarchise l’exécution des calculs et n’ont pas de volet Office. Ces fonctions utilisent un modèle de programmation légèrement différent et sont appelées fonctions sans interface utilisateur.
