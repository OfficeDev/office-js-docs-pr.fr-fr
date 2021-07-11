L’API JavaScript pour Office inclut deux modèles distincts :

- Les API **propres à l’application** fournissent des objets fortement typés qui peuvent être utilisés pour interagir avec des objets natifs d’une application Office spécifique. Par exemple, vous pouvez utiliser les API JavaScript pour Excel pour accéder à des feuilles de calcul, plages, tableaux, graphiques, etc. Les API spécifiques à l’application sont actuellement disponibles pour les applications Office suivantes.

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)
    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
    - [PowerPoint](../reference/overview/powerpoint-add-ins-reference-overview.md)
    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    Ce modèle d’API utilise [des promesses](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) et vous permet de spécifier plusieurs opérations dans chaque demande que vous envoyez à l’application Office. Les opérations de traitement par lots de cette manière peuvent améliorer sensiblement les performances des compléments dans les applications Office sur le Web. Les API propres à l’application ont été introduites avec Office 2016 et ne peuvent pas être utilisées pour interagir avec Office 2013.

    > [!NOTE]
    > Il existe également une API propre à l’application pour [Visio](../reference/overview/visio-javascript-reference-overview.md), mais vous pouvez l’utiliser uniquement dans les pages SharePoint Online pour interagir avec les diagrammes Visio incorporés dans la page. Les compléments Web Office ne sont pas pris en charge dans Visio.

    Visitez [Utilisation du modèle API propre à l’application](../develop/application-specific-api-model.md) pour en savoir plus sur ce modèle d’API.

- Les API **Communes** peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office. Ce modèle d’API utilise des[rappels](https://developer.mozilla.org/docs/Glossary/Callback_function), qui vous permettent de spécifier une seule opération dans chaque demande envoyée à l’application Office. Les API communes ont été introduites avec Office 2013 et peuvent être utilisées pour interagir avec Office 2013 ou version ultérieure. Si vous souhaitez plus en savoir sur le modèle objet API commun, qui inclut des API pour l’interaction avec Outlook et PowerPoint, veuillez consulter [Modèle d’objet API JavaScript communes](../develop/office-javascript-api-object-model.md).

> [!NOTE]
> Certaines fonctions Excel personnalisées s’exécutent dans le cadre d’une exécution unique qui hiérarchise l’exécution des calculs et n’ont pas de volet Office. Ces fonctions utilisent un modèle de programmation légèrement différent et sont appelées fonctions sans interface utilisateur.
