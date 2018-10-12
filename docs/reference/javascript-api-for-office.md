# <a name="javascript-api-for-office"></a>Interface API JavaScript pour Office

L’interface API JavaScript pour Office vous permet de créer des applications Web qui interagissent avec les modèles objets dans des applications hôtes Office. Votre application référencera la bibliothèque office.js, qui est un chargeur de script. La bibliothèque office.js charge les modèles objets applicables à l’application Office qui exécute le complément. Vous pouvez utiliser les modèles objets JavaScript suivants :

- **API communes** - API qui ont été introduites dans **Office 2013**. Elles sont chargées pour **toutes les applications hôtes Office** et connecte votre complément application avec l’application cliente Office. Le modèle objet contient les API qui sont propres aux clients Office et qui s’appliquent à plusieurs applications hôtes clientes Office. Tout ce contenu est sous **API partagée**. 

  **Outlook** utilise également la syntaxe de l’API commune. Tout ce qui se trouve sous l’alias Office contient des objets que vous pouvez utiliser pour écrire des scripts qui interagissent avec le contenu des documents Office, des feuilles de travail, des présentations, des éléments de courrier et des projets à partir de vos compléments Office. Vous devez utiliser ces API communes si votre complément cible Office 2013 et les versions ultérieures. Ce modèle objet utilise les rappels.

- **API propres aux hôtes** - API qui ont été introduites avec **Office 2016**. Ce modèle objet fournit des objets fortement typés propres aux hôtes qui correspondent aux objets familiers que vous voyez lorsque vous utilisez des clients Office et représente l’avenir des API JavaScript Office. Les API propres aux hôtes incluent actuellement les API JavaScript de Word et Excel.

## <a name="supported-host-applications"></a>Applications hôtes prises en charge

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [API partagée](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint et Project](requirement-sets/powerpoint-and-project-note.md) prennent en charge des compléments créés avec l’API JavaScript. Toutefois, ils n’ont actuellement pas d’API propre aux hôtes. Vous interagissez avec ces hôtes par le biais de l’API partagée.

En savoir plus sur les [hôtes pris en charge et les autres conditions requises](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins)

## <a name="open-api-specifications"></a>Spécifications d’API ouvertes

Au fur et à mesure que nous concevons et développons de nouvelles API pour les compléments Office, nous les mettons à votre disposition sur notre page de [spécifications d’API ouvertes](openspec.md) pour que vous puissiez fournir vos commentaires. Découvrez les nouvelles fonctionnalités dans le pipeline et donnez votre avis sur nos spécifications de conception.

## <a name="see-also"></a>Voir aussi

- [Référence de l’API JavaScript Office](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)