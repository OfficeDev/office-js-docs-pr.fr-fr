---
title: Limites des ressources et optimisation des performances pour les compléments Office
description: Découvrez les limites de ressources de la plateforme de complément Office, y compris le processeur et la mémoire.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 8c64e5a836d6b998ccd7022e71f595bb331bba8c
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293330"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Limites des ressources et optimisation des performances pour les compléments Office

Afin d’offrir la meilleure expérience utilisateur, assurez-vous que votre complément Office fonctionne dans les limites prévues en matière d’utilisation du cœur du processeur et de la mémoire, ainsi qu’en matière de fiabilité et, pour les compléments Outlook, de temps de réponse lors de l’évaluation des expressions régulières. Ces limites propres à l’utilisation des ressources d’exécution s’appliquent aux compléments exécutés sur des clients Office sous Windows et OS X mais pas sur des applications mobiles, ni dans un navigateur.

Par ailleurs, cette rubrique suggère des techniques de conception et d’implémentation de complément qui permettent de mieux contrôler les performances des compléments sur ordinateurs de bureau et périphériques mobiles.

## <a name="resource-usage-limits-for-add-ins"></a>Limites d’utilisation des ressources pour les compléments

Les limites d’utilisation des ressources d’exécution s’appliquent à tous les types de compléments Office. Ces limites contribuent à garantir les performances pour vos utilisateurs et à limiter les attaques par déni de service. Veillez à tester votre complément Office sur votre application Office cible à l’aide d’une plage de données possibles et mesurer ses performances par rapport aux limites d’utilisation au moment de l’exécution suivantes :

- **Utilisation du cœur du processeur** - Seuil d’utilisation d’un seul cœur de processeur de 90 %, observé à trois reprises dans des intervalles par défaut de 5 secondes.

   L’intervalle par défaut d’un client Office pour vérifier l’utilisation de l’UC est toutes les 5 secondes. Si le client Office détecte que l’utilisation de l’UC d’un complément est supérieure à la valeur de seuil, il affiche un message qui demande à l’utilisateur s’il souhaite continuer à exécuter le complément. Si l’utilisateur choisit de continuer, le client Office ne demande pas de nouveau à l’utilisateur pendant la session de modification. Les administrateurs peuvent souhaiter utiliser la clé de registre **AlertInterval** pour augmenter le seuil et réduire l’affichage de ce message d’avertissement si les utilisateurs exécutent des compléments faisant appel au processeur de manière intensive.

- **Utilisation de la mémoire** - Seuil d’utilisation de mémoire par défaut, qui est déterminé de manière dynamique en fonction de la mémoire physique disponible de l’appareil.

   Par défaut, lorsqu’un client Office détecte que l’utilisation de la mémoire physique sur un appareil dépasse 80% de la mémoire disponible, le client commence à surveiller l’utilisation de la mémoire du complément, au niveau du document pour les compléments de contenu et du volet Office, et au niveau de la boîte aux lettres pour les compléments Outlook. À un intervalle de 5 secondes par défaut, le client avertit l’utilisateur si l’utilisation de la mémoire physique pour un ensemble de compléments au niveau du document ou de la boîte aux lettres dépasse 50%. Cette limite d’utilisation de la mémoire est basée sur la mémoire physique plutôt que sur la mémoire virtuelle afin de garantir de bonnes performances sur les appareils disposant d’une mémoire vive (RAM) limitée, par exemple les tablettes. Les administrateurs peuvent remplacer ce paramètre dynamique par une limite explicite à l’aide de la clé de Registre Windows **MemoryAlertThreshold** en tant que paramètre global, IR régler l’intervalle d’alerte en utilisant la clé **AlertInterval** comme paramètre global.

- **Tolérance d’incident** - Limite par défaut de 4 incidents pour un complément.

   Les administrateurs peuvent ajuster le seuil relatif aux incidents en utilisant la clé de registre **RestartManagerRetryLimit**.

- **Blocage d’application** - Limitation à 5 secondes du seuil de blocage prolongé d’un complément.

   Cela a une incidence sur les expériences utilisateur du complément et de l’application Office. Lorsque cela se produit, l’application Office redémarre automatiquement tous les compléments actifs pour un document ou une boîte aux lettres (le cas échéant), et avertit l’utilisateur de l’absence de réponse du complément. Les compléments peuvent atteindre ce seuil lorsqu’ils ne cèdent pas régulièrement le traitement lors de l’exécution de tâches longues. Il existe des techniques permettant de garantir qu’aucun blocage ne se produira. Les administrateurs ne peuvent pas remplacer ce seuil.

### <a name="outlook-add-ins"></a>Compléments Outlook

§LTA Si un complément Outlook dépasse les seuils précédents en matière d’utilisation du cœur du processeur ou de la mémoire, ou en matière de tolérance d’incident, Outlook désactive le complément. Le Centre d’administration Exchange indique que l’état de l’application est désactivé.

> [!NOTE]
> Même si seuls les clients enrichis Outlook et non les clients non-Outlook sur le web ou les appareils mobiles contrôlent l’utilisation des ressources, si un client enrichi désactive un complément Outlook, ce complément est également désactivé pour une utilisation dans Outlook sur le web et les appareils mobiles.

En plus du cœur du processeur, de la mémoire et des règles de fiabilité, les compléments Outlook doivent respecter les règles suivantes lors de l’activation :

- **Temps de réponse des expressions régulières** - Seuil par défaut de 1 000 millisecondes pour Outlook afin d’évaluer toutes les expressions régulières contenues dans le manifeste d’un complément Outlook. Le dépassement du seuil oblige Outlook à retenter l’évaluation un peu plus tard.

    À l’aide d’une stratégie de groupe ou d’un paramètre spécifique de l’application dans le Registre Windows, les administrateurs peuvent ajuster cette valeur seuil par défaut de 1 000 millisecondes dans le paramètre **OutlookActivationAlertThreshold**.

- **Réévaluation des expressions régulières** -une limite par défaut de trois fois pour qu’Outlook réévalue toutes les expressions régulières dans un manifeste. Si l’évaluation échoue toutes les trois fois en dépassant le seuil applicable (soit la valeur par défaut de 1 000 millisecondes, soit une valeur spécifiée par **OutlookActivationAlertThreshold**, si ce paramètre existe dans le Registre Windows), Outlook désactive le complément Outlook. Le centre d’administration Exchange affiche l’état désactivé et le complément est désactivé pour une utilisation dans les clients riches Outlook, et Outlook sur le Web et les appareils mobiles.

    À l’aide d’une stratégie de groupe ou d’un paramètre spécifique de l’application dans le Registre Windows, les administrateurs peuvent ajuster ce nombre de tentatives d’évaluation dans le paramètre **OutlookActivationManagerRetryLimit**.

### <a name="excel-add-ins"></a>Compléments Excel

Si vous créez un complément Excel, tenez compte des limitations de taille suivantes lors de l’interaction avec le classeur :

- Excel sur le web a une limite de taille de charge utile de 5 Mo pour les demandes et les réponses. L’erreur `RichAPI.Error` est déclenchée en cas de dépassement de cette limite.
- Une plage est limitée à 5 millions cellules pour les opérations Get.

Si vous prévoyez que l’entrée de l’utilisateur dépasse ces limites, veillez à vérifier les données avant d’appeler `context.sync()` . Fractionnez l’opération en plusieurs parties si nécessaire. Veillez à appeler `context.sync()` pour chaque sous-opération afin d’éviter que ces opérations soient regroupées par lots.

Ces limitations sont généralement dépassées par les grandes plages. Votre complément peut utiliser [RangeAreas](/javascript/api/excel/excel.rangeareas) pour mettre à jour les cellules dans une plage plus grande de manière stratégique. Pour plus d’informations, consultez [travailler simultanément avec plusieurs plages dans des compléments Excel](../excel/excel-add-ins-multiple-ranges.md) .

### <a name="task-pane-and-content-add-ins"></a>Compléments de volet Office et de contenu

Si un complément de contenu ou de volet de tâches dépasse les seuils précédents sur le cœur de l’UC ou l’utilisation de la mémoire, ou la limite de tolérance pour les incidents, l’application Office correspondante affiche un avertissement pour l’utilisateur. À ce stade, l’utilisateur peut effectuer l’une des actions suivantes :

- Redémarrer le complément.
- Annuler les alertes supplémentaires de dépassement de seuil. Dans l’idéal, l’utilisateur devrait supprimer le complément du document. La poursuite de l’exécution du complément risquerait d’entraîner des problèmes supplémentaires au niveau des performances et de la stabilité.  

## <a name="verifying-resource-usage-issues-in-the-telemetry-log"></a>Vérification des problèmes d’utilisation des ressources dans le journal de télémétrie

Office propose un journal de télémétrie qui tient à jour un enregistrement de certains événements (chargement, ouverture, fermeture et erreurs) des solutions Office qui s’exécutent sur l’ordinateur local, notamment les problèmes d’utilisation des ressources dans une Complément Office. Si le journal de télémétrie est configuré, vous pouvez utiliser Excel pour l’ouvrir à partir de l’emplacement par défaut suivant sur votre disque local :

`%Users%\<Current user>\AppData\Local\Microsoft\Office\15.0\Telemetry`

Le journal de télémétrie comprend pour chaque événement suivi pour un complément, les date/heure de l’occurrence, l’ID d’événement, la gravité et une courte description de l’événement, le nom convivial et l’ID unique du complément, ainsi que l’application qui a enregistré l’événement. Vous pouvez actualiser le journal de télémétrie pour afficher les événements suivis. Le tableau suivant répertorie des exemples de compléments Outlook qui ont été suivis dans le journal de télémétrie.

|**Date/Heure**|**ID d’évènement**|**Gravité**|**Titre**|**Fichier**|**ID**|**Application**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|08/10/2012 17:57:10|7 ||Le manifeste du complément a été correctement téléchargé|Who’s Who|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|08/10/2012 17:57:01|7 ||Le manifeste du complément a été correctement téléchargé|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|

Le tableau suivant répertorie les événements que le journal de télémétrie suit pour les Compléments Office en général.

|**ID d’évènement**|**Titre**|**Gravité**|**Description**|
|:-----|:-----|:-----|:-----|
|7 |Le manifeste du complément a été correctement téléchargé||Le manifeste du complément Office a été chargé et lu avec succès par l’application Office.|
|8 |Échec du téléchargement du manifeste du complément|Critique|L’application Office n’a pas pu charger le fichier manifeste pour le complément Office à partir du catalogue SharePoint, du catalogue d’entreprise ou de AppSource.|
|9 |Impossible d’analyser le balisage du complément|Critique|L’application Office a chargé le manifeste du complément Office, mais n’a pas pu lire le balisage HTML de l’application.|
|10 |Le complément a trop sollicité le processeur|Critique|L’Complément Office a utilisé plus de 90 % des ressources du processeur sur une période de temps définie.|
|15 |Le complément a été désactivé en raison de l’expiration de la recherche de chaîne||§LTA Les compléments Outlook recherchent la ligne d’objet et le corps du message d’un courrier électronique pour déterminer s’ils doivent être affichés avec une expression régulière. Le complément Outlook répertorié dans la colonne **Fichier** a été désactivé par Outlook, car il a expiré à plusieurs reprises lors d’une tentative de mise en correspondance d’une expression régulière.|
|18 |Le complément a été fermé||L’application Office a pu fermer le complément Office.|
|neuf|Le complément a rencontré une erreur d’exécution|Critique|L'Complément Office a rencontré un problème qui l'a empêchée de s'exécuter. Pour plus de détails, consultez le journal **Alertes Microsoft Office** à l’aide de l’Observateur d’événements Windows sur l’ordinateur sur lequel l’erreur s’est produite.|
|vingtaine|Le complément n’a pas pu vérifier la licence|Critique|Les informations de licence de l'Complément Office n'ont pas pu être vérifiées et la licence a peut-être expiré. Pour plus de détails, consultez le journal **Alertes Microsoft Office** à l’aide de l’Observateur d’événements Windows sur l’ordinateur sur lequel l’erreur s’est produite.|

Pour plus d’informations, consultez [Déployer le Tableau de bord de télémétrie](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15)) et [Dépannage des fichiers et des solutions personnalisées d’Office avec le journal de télémétrie](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log)

## <a name="design-and-implementation-techniques"></a>Techniques de conception et d’implémentation

Bien que les limites en matière d’utilisation des ressources de l’UC et de la mémoire, de tolérance d’incident et de réactivité de l’interface utilisateur s’appliquent aux Compléments Office qui s’exécutent uniquement sur les clients enrichis, l’optimisation de l’utilisation de ces ressources et de la batterie doit constituer une priorité si vous voulez que votre complément s’exécute de manière satisfaisante sur tous les clients et appareils de prise en charge. L’optimisation est particulièrement importante si votre complément effectue des opérations de longue durée ou manipule de grands jeux de données. La liste suivante suggère certaines techniques permettant de fractionner les opérations gourmandes en ressources de l’UC en segments plus petits de sorte que votre complément puisse éviter une consommation excessive de ressources et que l’application Office puisse rester réactive :

- Dans un scénario où votre complément a besoin de lire un important volume de données à partir d’un jeu de données illimité, vous pouvez appliquer la pagination lors de la lecture des données dans une table ou réduire la taille des données à chaque opération de lecture raccourcie, plutôt que de tenter de terminer la lecture en une seule opération. Pour limiter la durée de l’entrée et de la sortie, vous pouvez utiliser la méthode [setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) de l’objet global. It also handles the data in defined chunks instead of randomly unbounded data. Une autre option consiste à utiliser [Async](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/async_function) pour gérer vos promesses.

- Si votre complément utilise un algorithme qui sollicite l’UC de manière intensive pour traiter un important volume de données, vous pouvez recourir aux API Web Worker afin d’effectuer une tâche de longue durée en arrière-plan pendant qu’un script distinct s’exécute au premier plan (par exemple, l’affichage de la progression d’une opération dans l’interface utilisateur). Les API Web Worker ne bloquent pas les activités des utilisateurs. En outre, elles permettent à la page HTML de rester réactive. Pour obtenir un exemple d’API Web Worker, voir les [bases des API Web Worker](https://www.html5rocks.com/tutorials/workers/basics/). Pour plus d’informations sur l’API Web Worker, voir [API Web Worker](https://developer.mozilla.org/docs/Web/API/Web_Workers_API).

- Si votre complément utilise un algorithme qui sollicite l’UC de manière intensive et si vous pouvez décomposer les entrées ou sorties de données en jeux de données de plus petite taille, créez un service web afin de lui passer les données et d’alléger la charge de l’UC, puis attendez un rappel asynchrone.

- Veillez à tester votre complément par rapport au volume de données le plus important possible, puis limitez votre complément pour lui permettre d’atteindre cette limite.

### <a name="performance-improvements-with-the-application-specific-apis"></a>Améliorations des performances avec les API spécifiques de l’application

Les conseils en matière de performances de l' [utilisation du modèle d’API propres aux applications](../develop/application-specific-api-model.md) fournissent des instructions sur l’utilisation des API propres aux applications pour Excel, OneNote, Visio et Word. En résumé, vous devez :

- [Chargez uniquement les propriétés nécessaires](../develop/application-specific-api-model.md#calling-load-without-parameters-not-recommended).
- [Réduisez le nombre d’appels de synchronisation ()](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-sync-calls). Pour plus d’informations sur la gestion des appels dans votre code, lisez [la rubrique Évitez d’utiliser la méthode Context. Sync dans des boucles](correlated-objects-pattern.md) `sync` .
- [Réduisez le nombre d’objets proxy créés](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-proxy-objects-created). Vous pouvez également déseffectuer le suivi des objets proxy, comme décrit dans la section suivante.

#### <a name="untrack-unneeded-proxy-objects"></a>Non-suivi des objets proxy inutiles

Les [objets proxy](../develop/application-specific-api-model.md#proxy-objects) persistent dans la mémoire jusqu’à ce qu' `RequestContext.sync()` il soit appelé. Les opérations par lots volumineux peuvent générer un grand nombre d’objets proxy qui sont uniquement utiles une fois pour le complément et peuvent être publiés à partir de la mémoire avant l’exécution du lot.

La `untrack()` méthode libère l’objet de la mémoire. Cette méthode est implémentée sur un grand nombre d’objets de proxy d’API propres à l’application. `untrack()`L’appel une fois votre complément effectué avec l’objet doit produire un gain de performances notable lors de l’utilisation d’un grand nombre d’objets proxy.

> [!NOTE]
> `Range.untrack()` est un raccourci pour [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-). N’importe quel objet proxy peut être non suivi en le supprimant de la liste d’objets suivis dans le contexte.

L’exemple de code Excel suivant remplit une plage sélectionnée avec des données, une cellule à la fois. Une fois que la valeur est ajoutée à la cellule, la plage représentant cette cellule est non suivie. Exécuter tout d’abord ce code avec une plage sélectionnée de 10 000 à 20 000 cellules, avec la `cell.untrack()` ligne et puis sans. Vous devez remarquer que le code est exécuté plus rapidement avec la `cell.untrack()` ligne que sans elle. Vous pouvez également remarquer un temps de réponse plus rapide par la suite, étant donné que l’étape de nettoyage prend moins de temps.

```js
Excel.run(async (context) => {
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (var i = 0; i < largeRange.rowCount; i++) {
        for (var j = 0; j < largeRange.columnCount; j++) {
            var cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // Call untrack() to release the range from memory.
            cell.untrack();
        }
    }

    await context.sync();
});
```

Notez que le fait de ne pas suivre les objets ne devient important que lorsque vous avez affaire à des milliers d’entre eux. La plupart des compléments n’ont pas besoin de gérer le suivi de l’objet proxy.

## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Optimisation des performances à l’aide de l’API JavaScript d’Excel](../excel/performance.md)
