# <a name="define-add-in-commands-in-your-manifest"></a>Définir des commandes de complément dans votre manifeste

Les commandes de complément sont un moyen de personnaliser facilement l’interface utilisateur d’Office par défaut en y ajoutant des éléments d’interface qui exécutent des actions, tels que des boutons personnalisés ajoutés au ruban. Pour créer des commandes, ajoutez un nœud **[VersionOverrides](../../reference/manifest/versionoverrides.md)** à un manifeste existant. 

Lorsqu’un manifeste contient l’élément **VersionOverrides**, les versions de Word, Excel, Outlook et PowerPoint prenant en charge les commandes de complément utiliseront les informations de cet élément pour charger le complément. Les versions antérieures des produits Office qui ne prennent pas en charge les commandes de complément ignoreront l’élément.

Lorsque les applications clientes reconnaissent le nœud **VersionOverrides**, le nom du complément s’affiche dans le ruban, et non dans un volet Office ou un volet de lecture/composition. Le complément n’apparaîtra pas dans les deux emplacements.
 
## <a name="versionoverrides"></a>VersionOverrides

L’élément [VersionOverrides](../../reference/manifest/versionoverrides.md) est l’élément racine qui contient des informations pour les commandes de complément implémentées par le complément. Il est pris en charge dans la version 1.1 et les versions ultérieures du schéma de manifeste.

Il existe deux versions du schéma **VersionOverrides**.

| Version du schéma | Description |
|----------------|-------------|
| 1.0 | Prend en charge les commandes de complément pour les versions de bureau des applications Office. | 
| 1.1 | Ajoute la prise en charge des [volets des tâches épinglables](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane) et des compléments mobiles. **Remarque :** ils sont actuellement uniquement pris en charge par Outlook 2016 pour Windows et Outlook pour iOS. |

Un complément peut prendre en charge plusieurs versions du schéma **VersionOverrides** en imbriquant des versions plus récentes à l’intérieur de la version précédente. Cela permet aux clients prendre en charge les versions plus récentes pour tirer parti des nouvelles fonctionnalités, tout en permettant aux clients plus anciens de charger la version plus ancienne. Voir la section sur la [mise en œuvre de plusieurs versions](../../reference/manifest/versionoverrides.md#implementing-multiple-versions) pour plus d’informations.

L’élément **VersionOverrides** inclut les éléments enfants suivants :

- [Description](../../reference/manifest/description.md)
- [Requirements](../../reference/manifest/requirements.md)
- [Hosts](../../reference/manifest/hosts.md)
- [Ressources](../../reference/manifest/resources.md)
- [VersionOverrides](../../reference/manifest/versionoverrides.md)

Le diagramme suivant illustre la hiérarchie des éléments utilisés pour définir des commandes de complément. 

![Hiérarchie des éléments de commandes de complément dans le manifeste](../images/080da303-51c4-4882-b74a-7ba11517c0ad.png)

## <a name="sample-manifests"></a>Exemple de manifestes

Pour un exemple de manifeste qui implémente les commandes de complément pour Word, Excel et PowerPoint, voir l’article sur l’[exemple de commandes de complément simples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Simple).

Pour un exemple de manifeste qui implémente des commandes de complément pour Outlook, voir l’article sur l’[exemple de fichier de manifeste pour un complément Outlook](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).

## <a name="additional-resources"></a>Ressources supplémentaires

- [Commandes de complément pour Outlook](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)
    
- [Manifestes des compléments Outlook](https://docs.microsoft.com/outlook/add-ins/manifests)
    
- [Démonstration de la commande du complément Outlook](https://github.com/OfficeDev/outlook-add-in-command-demo)
