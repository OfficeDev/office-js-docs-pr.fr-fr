# <a name="excel-add-ins-overview"></a>Présentation des compléments Excel

Un complément Excel vous permet d’étendre les fonctionnalités de l’application Excel sur plusieurs plateformes, notamment Office pour Windows, Office Online, Office pour Mac et Office pour iPad. Utilisez des compléments Excel dans un classeur aux fins suivantes :

- Interagir avec des objets Excel, lire et écrire des données Excel 
- Étendre les fonctionnalités à l’aide du volet Office web ou du volet de contenu 
- Ajouter des boutons personnalisés au ruban ou des éléments au menu contextuel
- Fournir une interaction améliorée à l’aide de la fenêtre de dialogue 

La plateforme de compléments Office fournit la structure et les API JavaScript Office.js qui vous permettent de créer et d’exécuter des compléments Excel. En utilisant la plateforme de compléments Office pour créer votre complément Excel, vous bénéficierez des avantages suivants :

* **Prise en charge sur plusieurs plateformes** : les compléments Excel s’exécutent dans Office pour Windows, Mac, iOS et Office Online.
* **Déploiement centralisé** : les administrateurs peuvent rapidement et facilement déployer des compléments Excel vers les utilisateurs d’une organisation.
* **Authentification unique (SSO)** : intégrez facilement votre complément Excel à l’aide de Microsoft Graph.
* **Utilisation des technologies web standard** : créez votre complément Excel en utilisant des technologies web connues telles qu’HTML, CSS et JavaScript.
* **Distribution via l’Office Store** : partagez votre complément Excel avec un public plus large en le publiant sur l’[Office Store](https://store.office.com/en-us/appshome.aspx).

> **Remarque** : Les compléments Excel sont différents des compléments COM ou VST, qui sont des solutions d’intégration Office antérieures s’exécutant uniquement sur Office pour Windows. Contrairement aux compléments COM, les compléments Excel ne nécessitent pas l’installation de code sur l’appareil d’un utilisateur ou dans Excel. 

## <a name="components-of-an-excel-add-in"></a>Composants d’un complément Excel 

Un complément Excel comprend deux composants de base : une application web et un fichier de configuration, appelé fichier manifeste. 

L’application web utilise l’[API JavaScript pour Office](../../reference/javascript-api-for-office.md) pour interagir avec des objets dans Excel et peut également faciliter l’interaction avec les ressources en ligne. Par exemple, un complément peut effectuer une des opérations suivantes :

* Créer, lire, mettre à jour et supprimer des données dans le classeur (feuilles de calcul, plages, tableaux, graphiques, éléments nommés, etc.).
* Effectuer une autorisation utilisateur avec un service en ligne à l’aide du flux OAuth 2.0 standard.
* Émettre des demandes d’API à Microsoft Graph ou toute autre API.

L’application web peut être hébergée sur un serveur web et peut être créée à l’aide de structures de côté client (par exemple, Angular, React, jQuery) ou des technologies côté serveur (par exemple, ASP.NET, Node.js, PHP).

Le [manifeste](../overview/add-in-manifests.md) est un fichier de configuration XML qui définit la façon dont le complément est intégré dans les clients Office en spécifiant des paramètres et fonctionnalités telles que : 

* L’URL de l’application web du complément.
* Le nom d’affichage, la description, l’ID, la version et les paramètres régionaux par défaut du complément.
* La manière dont le complément est intégré à Excel, y compris toute interface utilisateur personnalisée créée par le complément (boutons du ruban, menus contextuels, etc.).
* Les autorisations requise par le complément, comme la lecture du document ou l’écriture dans celui-ci.

Pour permettre aux utilisateurs finals d’installer et d’utiliser un complément Excel, vous devez publier son manifeste dans l’Office Store ou dans un catalogue de compléments. 

## <a name="capabilities-of-an-excel-add-in"></a>Fonctionnalités d’un complément Excel

En plus d’interagir avec le contenu du classeur, les compléments Excel peuvent ajouter des boutons personnalisés au ruban ou des commandes de menu, insérer des volets de tâches, ouvrir des boîtes de dialogue et même incorporer des objets riches web, tels que des graphiques ou des visualisations interactives dans une feuille de calcul, comme indiqué dans les captures d’écran ci-dessous. Pour plus d’informations sur chacune de ces fonctionnalités, consultez [Étendre les fonctionnalités d’Excel](excel-add-ins-extend-excel.md).

**Boutons personnalisés du ruban**

![Commandes de complément](../../images/Excel_add-in_commands_Script-Lab.png)

**Volet Office**

![Volet Office pour le complément](../../images/Excel_add-in_task_pane_Insights.png)

**Boîte de dialogue**

![Boîte de dialogue de complément](../../images/Excel_add-in_dialog_choose-number.png)

**Complément de contenu**

![complément de contenu](../../images/Excel_add-in_content_map.png)

## <a name="javascript-apis-to-interact-with-workbook-content"></a>API JavaScript permettant d’interagir avec le contenu du classeur

Un complément Excel interagit avec des objets dans Excel en utilisant l’[API JavaScript pour Office](../../reference/javascript-api-for-office.md), qui inclut deux modèles d’objets JavaScript :

* **API JavaScript pour Excel** : incluse dans Office 2016, l’[API JavaScript pour Excel](../../reference/excel/excel-add-ins-reference-overview.md) fournit des objets Excel fortement typés que vous pouvez utiliser pour accéder aux feuilles de calcul, aux plages, aux tableaux, aux graphiques et bien plus encore. 

* **API partagée** : incluse dans Office 2013, l’API partagée vous permet d’accéder à des fonctionnalités, comme l’interface utilisateur, les boîtes de dialogue et les paramètres du client, qui sont communes à plusieurs types d’applications hôtes, telles que Word, Excel et PowerPoint. Étant donné que l’API partagée fournit des fonctionnalités limitées pour une interaction avec Excel, vous pouvez l’utiliser si votre complément doit s’exécuter sur Excel 2013.

## <a name="next-steps"></a>Étapes suivantes

Apprenez à [créer votre premier complément Excel](excel-add-ins-get-started-overview.md). Découvrez ensuite les [concepts fondamentaux](excel-add-ins-core-concepts.md) de la création de compléments Excel.

## <a name="additional-resources"></a>Ressources supplémentaires

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Meilleures pratiques en matière de développement de compléments Office](../overview/add-in-development-best-practices.md)
- [Instructions de conception pour les compléments Office](../design/add-in-design.md)
- [Concepts de base de l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Référence de l’API JavaScript pour Excel](../../reference/excel/excel-add-ins-reference-overview.md)
