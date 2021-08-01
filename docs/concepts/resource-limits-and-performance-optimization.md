---
title: Limites des ressources et optimisation des performances pour les compléments Office
description: Découvrez les limites de ressources de la plateforme de Office, y compris le processeur et la mémoire.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 750f10880249a9c9a8720f870f4bc5ea4e576e8e
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671274"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Limites des ressources et optimisation des performances pour les compléments Office

Afin d’offrir la meilleure expérience utilisateur, assurez-vous que votre complément Office fonctionne dans les limites prévues en matière d’utilisation du cœur du processeur et de la mémoire, ainsi qu’en matière de fiabilité et, pour les compléments Outlook, de temps de réponse lors de l’évaluation des expressions régulières. Ces limites propres à l’utilisation des ressources d’exécution s’appliquent aux compléments exécutés sur des clients Office sous Windows et OS X mais pas sur des applications mobiles, ni dans un navigateur.

Par ailleurs, cette rubrique suggère des techniques de conception et d’implémentation de complément qui permettent de mieux contrôler les performances des compléments sur ordinateurs de bureau et périphériques mobiles.

## <a name="resource-usage-limits-for-add-ins"></a>Limites d’utilisation des ressources pour les compléments

Les limites d’utilisation des ressources au moment de l’Office s’appliquent à tous les types de Office des applications. Ces limites permettent de garantir les performances de vos utilisateurs et d’atténuer les attaques par déni de service. N’oubliez pas de tester votre Office sur votre application Office cible à l’aide d’une plage de données possibles et de mesurer ses performances par rapport aux limites d’utilisation au moment de l’exécution suivantes.

- **Utilisation du cœur du processeur** - Seuil d’utilisation d’un seul cœur de processeur de 90 %, observé à trois reprises dans des intervalles par défaut de 5 secondes.

   L’intervalle par défaut pour qu Office client vérifie l’utilisation du cœur du processeur toutes les 5 secondes. Si le client Office détecte que l’utilisation du cœur du processeur d’un module est supérieure à la valeur de seuil, il affiche un message demandant à l’utilisateur s’il souhaite continuer à l’exécution du module. Si l’utilisateur choisit de continuer, le client Office ne lui demande pas de nouveau pendant cette session de modification. Les administrateurs peuvent souhaiter utiliser la clé de registre **AlertInterval** pour augmenter le seuil et réduire l’affichage de ce message d’avertissement si les utilisateurs exécutent des compléments faisant appel au processeur de manière intensive.

- **Utilisation de la mémoire** - Seuil d’utilisation de mémoire par défaut, qui est déterminé de manière dynamique en fonction de la mémoire physique disponible de l’appareil.

   Par défaut, lorsqu’un client Office détecte que l’utilisation de la mémoire physique sur un appareil dépasse 80 % de la mémoire disponible, le client commence à surveiller l’utilisation de la mémoire du module, au niveau du document pour les applications de contenu et du volet Des tâches, et au niveau de la boîte aux lettres pour les applications Outlook. À un intervalle par défaut de 5 secondes, le client avertit l’utilisateur si l’utilisation de la mémoire physique pour un ensemble de modules au niveau du document ou de la boîte aux lettres dépasse 50 %. Cette limite d’utilisation de la mémoire est basée sur la mémoire physique plutôt que sur la mémoire virtuelle afin de garantir de bonnes performances sur les appareils disposant d’une mémoire vive (RAM) limitée, par exemple les tablettes. Les administrateurs peuvent remplacer ce paramètre dynamique par une limite explicite en utilisant la clé de Registre **MemoryAlertThreshold** Windows comme paramètre global, puis ajuster l’intervalle d’alerte en utilisant la clé **AlertInterval** comme paramètre global.

- **Tolérance d’incident** - Limite par défaut de 4 incidents pour un complément.

   Les administrateurs peuvent ajuster le seuil relatif aux incidents en utilisant la clé de registre **RestartManagerRetryLimit**.

- **Blocage d’application** - Limitation à 5 secondes du seuil de blocage prolongé d’un complément.

   Cela a une incidence sur les expériences de l’utilisateur du Office application. Lorsque cela se produit, l’application Office redémarre automatiquement tous les applications actives pour un document ou une boîte aux lettres (le cas échéant) et avertit l’utilisateur de l’indisponible du module. Les compléments peuvent atteindre ce seuil lorsqu’ils ne cèdent pas régulièrement le traitement lors de l’exécution de tâches longues. Il existe des techniques permettant de garantir qu’aucun blocage ne se produira. Les administrateurs ne peuvent pas remplacer ce seuil.

### <a name="outlook-add-ins"></a>Compléments Outlook

§LTA Si un complément Outlook dépasse les seuils précédents en matière d’utilisation du cœur du processeur ou de la mémoire, ou en matière de tolérance d’incident, Outlook désactive le complément. Le Centre d’administration Exchange indique que l’état de l’application est désactivé.

> [!NOTE]
> Même si seuls les clients enrichis Outlook et non les clients non-Outlook sur le web ou les appareils mobiles contrôlent l’utilisation des ressources, si un client enrichi désactive un complément Outlook, ce complément est également désactivé pour une utilisation dans Outlook sur le web et les appareils mobiles.

Outre les règles relatives au cœur de l’UC, à la mémoire et à la fiabilité, Outlook compléments doivent respecter les règles suivantes lors de l’activation.

- **Temps de réponse des expressions régulières** - Seuil par défaut de 1 000 millisecondes pour Outlook afin d’évaluer toutes les expressions régulières contenues dans le manifeste d’un complément Outlook. Le dépassement du seuil oblige Outlook à retenter l’évaluation un peu plus tard.

    À l’aide d’une stratégie de groupe ou d’un paramètre spécifique de l’application dans le Registre Windows, les administrateurs peuvent ajuster cette valeur seuil par défaut de 1 000 millisecondes dans le paramètre **OutlookActivationAlertThreshold**.

- **Réévaluation des expressions régulières** : limite par défaut de trois fois pour Outlook réévaluer toutes les expressions régulières dans un manifeste. Si l’évaluation échoue à trois reprises en dépassant le seuil applicable (qui est la valeur par défaut de 1 000 millisecondes ou une valeur spécifiée par **OutlookActivationAlertThreshold,** si ce paramètre existe dans le Registre Windows), Outlook désactive le Outlook. Le Centre d’administration Exchange affiche l’état désactivé et le module est désactivé pour être utilisé dans les clients Outlook riches, ainsi que sur Outlook sur le web appareils mobiles et mobiles.

    À l’aide d’une stratégie de groupe ou d’un paramètre spécifique de l’application dans le Registre Windows, les administrateurs peuvent ajuster ce nombre de tentatives d’évaluation dans le paramètre **OutlookActivationManagerRetryLimit**.

### <a name="excel-add-ins"></a>Compléments Excel

Si vous construisez un Excel, n’ignorez pas les limitations de taille suivantes lors de l’interaction avec le workbook.

- Excel sur le web a une limite de taille de charge utile de 5 Mo pour les demandes et les réponses. L’erreur `RichAPI.Error` est déclenchée en cas de dépassement de cette limite.
- Une plage est limitée à cinq millions de cellules pour obtenir des opérations.

Si vous pensez que l’entrée utilisateur dépasse ces limites, veillez à vérifier les données avant `context.sync()` d’appeler. Fractionner l’opération en plus petites parties selon les besoins. N’oubliez pas d’appeler chaque sous-opération pour éviter que ces `context.sync()` opérations ne soient à nouveau rassemblées par lots.

Ces limitations sont généralement dépassées par de grandes plages. Votre add-in peut être en mesure d’utiliser [RangeAreas](/javascript/api/excel/excel.rangeareas) pour mettre à jour de manière stratégique des cellules dans une plage plus étendue. Pour [plus d’informations,](../excel/excel-add-ins-multiple-ranges.md) voir Travailler avec plusieurs plages simultanément dans Excel de recherche.

### <a name="task-pane-and-content-add-ins"></a>Compléments de volet Office et de contenu

Si un application de contenu ou du volet Des tâches dépasse les seuils précédents sur l’utilisation du cœur de l’UC ou de la mémoire, ou la limite de tolérance pour les incidents, l’application Office correspondante affiche un avertissement pour l’utilisateur. À ce stade, l’utilisateur peut effectuer l’une des actions suivantes :

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
|7 |Le manifeste du complément a été correctement téléchargé||Le manifeste du Office a été chargé et lu avec succès par l’application Office’application.|
|8 |Échec du téléchargement du manifeste du complément|Critique|L Office’application n’a pas pu charger le fichier manifeste du Office à partir du catalogue SharePoint, du catalogue d’entreprise ou d’AppSource.|
|9 |Impossible d’analyser le balisage du complément|Critique|L Office’application a chargé Office manifeste de l’application, mais n’a pas pu lire le code HTML de l’application.|
|10 |Le complément a trop sollicité le processeur|Critique|L’Complément Office a utilisé plus de 90 % des ressources du processeur sur une période de temps définie.|
|15|Le complément a été désactivé en raison de l’expiration de la recherche de chaîne||§LTA Les compléments Outlook recherchent la ligne d’objet et le corps du message d’un courrier électronique pour déterminer s’ils doivent être affichés avec une expression régulière. Le complément Outlook répertorié dans la colonne **Fichier** a été désactivé par Outlook, car il a expiré à plusieurs reprises lors d’une tentative de mise en correspondance d’une expression régulière.|
|18 |Le complément a été fermé||L Office’application a pu fermer le Office le module.|
|19|Le complément a rencontré une erreur d’exécution|Critique|L'Complément Office a rencontré un problème qui l'a empêchée de s'exécuter. Pour plus de détails, consultez le journal **Alertes Microsoft Office** à l’aide de l’Observateur d’événements Windows sur l’ordinateur sur lequel l’erreur s’est produite.|
|20|Le complément n’a pas pu vérifier la licence|Critique|Les informations de licence de l'Complément Office n'ont pas pu être vérifiées et la licence a peut-être expiré. Pour plus de détails, consultez le journal **Alertes Microsoft Office** à l’aide de l’Observateur d’événements Windows sur l’ordinateur sur lequel l’erreur s’est produite.|

Pour plus d’informations, consultez [Déployer le Tableau de bord de télémétrie](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15)) et [Dépannage des fichiers et des solutions personnalisées d’Office avec le journal de télémétrie](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log)

## <a name="design-and-implementation-techniques"></a>Techniques de conception et d’implémentation

Bien que les limites en matière d’utilisation des ressources de l’UC et de la mémoire, de tolérance d’incident et de réactivité de l’interface utilisateur s’appliquent aux Compléments Office qui s’exécutent uniquement sur les clients enrichis, l’optimisation de l’utilisation de ces ressources et de la batterie doit constituer une priorité si vous voulez que votre complément s’exécute de manière satisfaisante sur tous les clients et appareils de prise en charge. L’optimisation est particulièrement importante si votre complément effectue des opérations de longue durée ou manipule de grands jeux de données. La liste suivante suggère certaines techniques permettant de décomposer les opérations qui consomment beaucoup de processeurs ou de données en blocs plus petits afin que votre application puisse éviter une consommation excessive de ressources et que l’application Office puisse rester réactive.

- Dans un scénario où votre complément a besoin de lire un important volume de données à partir d’un jeu de données illimité, vous pouvez appliquer la pagination lors de la lecture des données dans une table ou réduire la taille des données à chaque opération de lecture raccourcie, plutôt que de tenter de terminer la lecture en une seule opération. Pour ce faire, vous pouvez utiliser la [méthode setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) de l’objet global pour limiter la durée de l’entrée et de la sortie. It also handles the data in defined chunks instead of randomly unbounded data. Une autre option consiste à utiliser [la async](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/async_function) pour gérer vos promesses.

- Si votre complément utilise un algorithme qui sollicite l’UC de manière intensive pour traiter un important volume de données, vous pouvez recourir aux API Web Worker afin d’effectuer une tâche de longue durée en arrière-plan pendant qu’un script distinct s’exécute au premier plan (par exemple, l’affichage de la progression d’une opération dans l’interface utilisateur). Les API Web Worker ne bloquent pas les activités des utilisateurs. En outre, elles permettent à la page HTML de rester réactive. Pour obtenir un exemple d’API Web Worker, voir les [bases des API Web Worker](https://www.html5rocks.com/tutorials/workers/basics/). Pour plus d’informations sur l’API Web Worker, voir [API Web Worker](https://developer.mozilla.org/docs/Web/API/Web_Workers_API).

- Si votre complément utilise un algorithme qui sollicite l’UC de manière intensive et si vous pouvez décomposer les entrées ou sorties de données en jeux de données de plus petite taille, créez un service web afin de lui passer les données et d’alléger la charge de l’UC, puis attendez un rappel asynchrone.

- Veillez à tester votre complément par rapport au volume de données le plus important possible, puis limitez votre complément pour lui permettre d’atteindre cette limite.

### <a name="performance-improvements-with-the-application-specific-apis"></a>Améliorations des performances avec les API propres à l’application

Les conseils de performances dans l’utilisation du modèle API propre à [l’application](../develop/application-specific-api-model.md) fournissent des conseils lors de l’utilisation des API propres à l’application pour Excel, OneNote, Visio et Word. En résumé, vous devez :

- [Chargez uniquement les propriétés nécessaires.](../develop/application-specific-api-model.md#calling-load-without-parameters-not-recommended)
- [Réduisez le nombre d’appels sync().](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-sync-calls) Pour [plus d’informations sur](correlated-objects-pattern.md) la gestion des appels dans votre code, évitez d’utiliser la méthode context.sync en `sync` boucle.
- [Réduisez le nombre d’objets proxy créés.](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-proxy-objects-created) Vous pouvez également désister les objets proxy, comme décrit dans la section suivante.

#### <a name="untrack-unneeded-proxy-objects"></a>Untrack unneeded proxy objects

[Les objets proxy sont persistants](../develop/application-specific-api-model.md#proxy-objects) en mémoire `RequestContext.sync()` jusqu’à ce qu’ils soient appelés. Les opérations par lots volumineux peuvent générer un grand nombre d’objets proxy qui sont uniquement utiles une fois pour le complément et peuvent être publiés à partir de la mémoire avant l’exécution du lot.

La `untrack()` méthode libère l’objet de la mémoire. Cette méthode est implémentée sur de nombreux objets proxy d’API propres à l’application. L’appel après la fin de votre ajout avec l’objet devrait produire un avantage de performances perceptible lors de l’utilisation d’un grand nombre `untrack()` d’objets proxy.

> [!NOTE]
> `Range.untrack()` est un raccourci pour [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove_object_). N’importe quel objet proxy peut être non suivi en le supprimant de la liste d’objets suivis dans le contexte.

L’exemple Excel code suivant remplit une plage sélectionnée avec des données, une cellule à la fois. Une fois que la valeur est ajoutée à la cellule, la plage représentant cette cellule est non suivie. Exécuter tout d’abord ce code avec une plage sélectionnée de 10 000 à 20 000 cellules, avec la `cell.untrack()` ligne et puis sans. Vous devez remarquer que le code est exécuté plus rapidement avec la `cell.untrack()` ligne que sans elle. Vous pouvez également remarquer un temps de réponse plus rapide par la suite, étant donné que l’étape de nettoyage prend moins de temps.

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

Notez que le fait d’avoir besoin d’un suivi des objets devient important uniquement lorsque vous avez affaire à des milliers d’entre eux. La plupart des add-ins n’ont pas besoin de gérer le suivi d’objets proxy.

## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Optimisation des performances à l’aide de l’API JavaScript d’Excel](../excel/performance.md)
