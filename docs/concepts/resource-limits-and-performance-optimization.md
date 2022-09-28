---
title: Limites des ressources et optimisation des performances pour les compléments Office
description: Découvrez les limites de ressources de la plateforme de complément Office, notamment le processeur et la mémoire.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8465eb654795b538182e01d33b2fc57ddb35eaa0
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092902"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Limites des ressources et optimisation des performances pour les compléments Office

To create the best experience for your users, ensure that your Office Add-in performs within specific limits for CPU core and memory usage, reliability, and, for Outlook add-ins, the response time for evaluating regular expressions. These run-time resource usage limits apply to add-ins running in Office clients on Windows and OS X, but not on mobile apps or in a browser.

Par ailleurs, cette rubrique suggère des techniques de conception et d’implémentation de complément qui permettent de mieux contrôler les performances des compléments sur ordinateurs de bureau et périphériques mobiles.

## <a name="resource-usage-limits-for-add-ins"></a>Limites d’utilisation des ressources pour les compléments

Les limites d’utilisation des ressources au moment de l’exécution s’appliquent à tous les types de compléments Office. Ces limites permettent de garantir les performances de vos utilisateurs et d’atténuer les attaques par déni de service. Veillez à tester votre complément Office sur votre application Office cible à l’aide d’une plage de données possibles et à mesurer ses performances par rapport aux limites d’utilisation au moment de l’exécution suivantes.

- **Utilisation du cœur du processeur** - Seuil d’utilisation d’un seul cœur de processeur de 90 %, observé à trois reprises dans des intervalles par défaut de 5 secondes.

   L’intervalle par défaut pour qu’un client Office vérifie l’utilisation du cœur du processeur est toutes les 5 secondes. Si le client Office détecte que l’utilisation principale du processeur d’un complément est supérieure à la valeur de seuil, il affiche un message demandant si l’utilisateur souhaite continuer à exécuter le complément. Si l’utilisateur choisit de continuer, le client Office ne le demande plus pendant cette session de modification. Les administrateurs peuvent souhaiter utiliser la clé de registre **AlertInterval** pour augmenter le seuil et réduire l’affichage de ce message d’avertissement si les utilisateurs exécutent des compléments faisant appel au processeur de manière intensive.

- **Utilisation de la mémoire** - Seuil d’utilisation de mémoire par défaut, qui est déterminé de manière dynamique en fonction de la mémoire physique disponible de l’appareil.

   Par défaut, lorsqu’un client Office détecte que l’utilisation de la mémoire physique sur un appareil dépasse 80 % de la mémoire disponible, le client commence à surveiller l’utilisation de la mémoire du complément, au niveau du document pour le contenu et les compléments du volet Office, et au niveau de la boîte aux lettres pour les compléments Outlook. À un intervalle par défaut de 5 secondes, le client avertit l’utilisateur si l’utilisation de la mémoire physique pour un ensemble de compléments au niveau du document ou de la boîte aux lettres dépasse 50 %. Cette limite d’utilisation de la mémoire est basée sur la mémoire physique plutôt que sur la mémoire virtuelle afin de garantir de bonnes performances sur les appareils disposant d’une mémoire vive (RAM) limitée, par exemple les tablettes. Les administrateurs peuvent remplacer ce paramètre dynamique par une limite explicite à l’aide de la clé de Registre Windows **MemoryAlertThreshold** comme paramètre global, et ajuster l’intervalle d’alerte à l’aide de la clé **AlertInterval** comme paramètre global.

- **Tolérance d’incident** - Limite par défaut de 4 incidents pour un complément.

   Les administrateurs peuvent ajuster le seuil relatif aux incidents en utilisant la clé de registre **RestartManagerRetryLimit**.

- **Blocage d’application** - Limitation à 5 secondes du seuil de blocage prolongé d’un complément.

   Cela affecte les expériences de l’utilisateur du complément et de l’application Office. Lorsque cela se produit, l’application Office redémarre automatiquement tous les compléments actifs d’un document ou d’une boîte aux lettres (le cas échéant) et avertit l’utilisateur de l’absence de réponse du complément. Les compléments peuvent atteindre ce seuil lorsqu’ils ne cèdent pas régulièrement le traitement lors de l’exécution de tâches longues. Il existe des techniques permettant de garantir qu’aucun blocage ne se produira. Les administrateurs ne peuvent pas remplacer ce seuil.

### <a name="outlook-add-ins"></a>Compléments Outlook

If any Outlook add-in exceeds the preceding thresholds for CPU core or memory usage, or tolerance limit for crashes, Outlook disables the add-in. The Exchange Admin Center displays the disabled status of the app.

> [!NOTE]
> Même si seuls les clients enrichis Outlook et non les clients non-Outlook sur le web ou les appareils mobiles contrôlent l’utilisation des ressources, si un client enrichi désactive un complément Outlook, ce complément est également désactivé pour une utilisation dans Outlook sur le web et les appareils mobiles.

Outre les règles de base, de mémoire et de fiabilité du processeur, les compléments Outlook doivent respecter les règles suivantes sur l’activation.

- **Regular expressions response time** - A default threshold of 1,000 milliseconds for Outlook to evaluate all regular expressions in the manifest of an Outlook add-in. Exceeding the threshold causes Outlook to retry evaluation at a later time.

    À l’aide d’une stratégie de groupe ou d’un paramètre spécifique de l’application dans le Registre Windows, les administrateurs peuvent ajuster cette valeur seuil par défaut de 1 000 millisecondes dans le paramètre **OutlookActivationAlertThreshold**.

- **Réévaluation des expressions régulières** : limite par défaut de trois fois pour Qu’Outlook réévalue toutes les expressions régulières dans un manifeste. Si l’évaluation échoue trois fois en dépassant le seuil applicable (qui est soit la valeur par défaut de 1 000 millisecondes, soit une valeur spécifiée par **OutlookActivationAlertThreshold**, si ce paramètre existe dans le Registre Windows), Outlook désactive le complément Outlook. Le Centre de Administration Exchange affiche l’état désactivé, et le complément est désactivé pour une utilisation dans les clients riches Outlook, ainsi que les Outlook sur le web et les appareils mobiles.

    À l’aide d’une stratégie de groupe ou d’un paramètre spécifique de l’application dans le Registre Windows, les administrateurs peuvent ajuster ce nombre de tentatives d’évaluation dans le paramètre **OutlookActivationManagerRetryLimit**.

### <a name="excel-add-ins"></a>Compléments Excel

Si vous créez un complément Excel, tenez compte des limitations de taille suivantes lors de l’interaction avec le classeur.

- Excel sur le web a une limite de taille de charge utile de 5 Mo pour les demandes et les réponses. L’erreur `RichAPI.Error` est déclenchée en cas de dépassement de cette limite.
- Une plage est limitée à cinq millions de cellules pour les opérations d’obtention.

Si vous vous attendez à ce que l’entrée utilisateur dépasse ces limites, veillez à vérifier les données avant d’appeler `context.sync()`. Fractionnez l’opération en petits morceaux selon les besoins. Veillez à appeler `context.sync()` chaque sous-opération pour éviter que ces opérations ne soient à nouveau regroupées.

Ces limitations sont généralement dépassées par de grandes plages. Votre complément peut être en mesure d’utiliser [RangeAreas](/javascript/api/excel/excel.rangeareas) pour mettre à jour stratégiquement les cellules dans une plage plus grande. Pour plus d’informations sur l’utilisation `RangeAreas`, consultez [Utiliser plusieurs plages simultanément dans les compléments Excel](../excel/excel-add-ins-multiple-ranges.md). Pour plus d’informations sur l’optimisation de la taille de charge utile dans Excel, consultez les [meilleures pratiques en matière de limite de taille de charge utile](../excel/performance.md#payload-size-limit-best-practices).

### <a name="task-pane-and-content-add-ins"></a>Compléments de volet Office et de contenu

Si un complément de contenu ou de volet Office dépasse les seuils précédents sur l’utilisation du cœur du processeur ou de la mémoire, ou la limite de tolérance pour les incidents, l’application Office correspondante affiche un avertissement pour l’utilisateur. À ce stade, l’utilisateur peut effectuer l’une des actions suivantes :

- Redémarrer le complément.
- Cancel further alerts about exceeding that threshold. Ideally, the user should then delete the add-in from the document; continuing the add-in would risk further performance and stability issues.  

## <a name="verify-resource-usage-issues-in-the-telemetry-log"></a>Vérifier les problèmes d’utilisation des ressources dans le journal de télémétrie

Office propose un journal de télémétrie qui tient à jour un enregistrement de certains événements (chargement, ouverture, fermeture et erreurs) des solutions Office qui s’exécutent sur l’ordinateur local, notamment les problèmes d’utilisation des ressources dans une Complément Office. Si vous avez configuré le journal de télémétrie, vous pouvez utiliser Excel pour ouvrir le journal de télémétrie à l’emplacement par défaut suivant sur votre lecteur local.

`%Users%\<Current user>\AppData\Local\Microsoft\Office\15.0\Telemetry`

For each event that the Telemetry Log tracks for an add-in, there is a date/time of the occurrence, event ID, severity, and short descriptive title for the event, the friendly name and unique ID of the add-in, and the application that logged the event. You can refresh the Telemetry Log to see the current tracked events. The following table shows examples of Outlook add-ins that were tracked in the Telemetry log.

|Date/Heure|ID d’événement|Severity|Titre|File|ID|Application|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|08/10/2012 17:57:10|7 |*Non applicable*|Le manifeste du complément a été correctement téléchargé|Who’s Who|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|08/10/2012 17:57:01|7 |*Non applicable*|Le manifeste du complément a été correctement téléchargé|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|

Le tableau suivant répertorie les événements que le journal de télémétrie suit pour les Compléments Office en général.

|ID d’événement|Titre|Severity|Description|
|:-----|:-----|:-----|:-----|
|7 |Le manifeste du complément a été correctement téléchargé|*Non applicable*|Le manifeste du complément Office a été correctement chargé et lu par l’application Office.|
|8 |Échec du téléchargement du manifeste du complément|Critique|L’application Office n’a pas pu charger le fichier manifeste du complément Office à partir du catalogue SharePoint, du catalogue d’entreprise ou d’AppSource.|
|9 |Impossible d’analyser le balisage du complément|Critique|L’application Office a chargé le manifeste du complément Office, mais n’a pas pu lire le balisage HTML de l’application.|
|10|Le complément a trop sollicité le processeur|Critique|L’Complément Office a utilisé plus de 90 % des ressources du processeur sur une période de temps définie.|
|15|Le complément a été désactivé en raison de l’expiration de la recherche de chaîne|*Non applicable*|Outlook add-ins search the subject line and message of an e-mail to determine whether they should be displayed by using a regular expression. The Outlook add-in listed in the **File** column was disabled by Outlook because it timed out repeatedly while trying to match a regular expression.|
|18 |Le complément a été fermé|*Non applicable*|L’application Office a pu fermer le complément Office avec succès.|
|19|Le complément a rencontré une erreur d’exécution|Critique|L'Complément Office a rencontré un problème qui l'a empêchée de s'exécuter. Pour plus de détails, consultez le journal **Alertes Microsoft Office** à l’aide de l’Observateur d’événements Windows sur l’ordinateur sur lequel l’erreur s’est produite.|
|20|Le complément n’a pas pu vérifier la licence|Critique|Les informations de licence de l'Complément Office n'ont pas pu être vérifiées et la licence a peut-être expiré. Pour plus de détails, consultez le journal **Alertes Microsoft Office** à l’aide de l’Observateur d’événements Windows sur l’ordinateur sur lequel l’erreur s’est produite.|

Pour plus d’informations, consultez [Déployer le Tableau de bord de télémétrie](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15)) et [Dépannage des fichiers et des solutions personnalisées d’Office avec le journal de télémétrie](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log)

## <a name="design-and-implementation-techniques"></a>Techniques de conception et d’implémentation

Bien que les limites en matière d’utilisation des ressources de l’UC et de la mémoire, de tolérance d’incident et de réactivité de l’interface utilisateur s’appliquent aux Compléments Office qui s’exécutent uniquement sur les clients enrichis, l’optimisation de l’utilisation de ces ressources et de la batterie doit constituer une priorité si vous voulez que votre complément s’exécute de manière satisfaisante sur tous les clients et appareils de prise en charge. L’optimisation est particulièrement importante si votre complément effectue des opérations de longue durée ou manipule de grands jeux de données. La liste suivante suggère certaines techniques pour décomposer les opérations gourmandes en ressources processeur ou gourmandes en données en blocs plus petits afin que votre complément puisse éviter une consommation excessive de ressources et que l’application Office puisse rester réactive.

- Dans un scénario où votre complément a besoin de lire un important volume de données à partir d’un jeu de données illimité, vous pouvez appliquer la pagination lors de la lecture des données dans une table ou réduire la taille des données à chaque opération de lecture raccourcie, plutôt que de tenter de terminer la lecture en une seule opération. Pour ce faire, vous pouvez utiliser la méthode [setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) de l’objet global pour limiter la durée d’entrée et de sortie. It also handles the data in defined chunks instead of randomly unbounded data. Une autre option consiste à utiliser [asynchrone](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/async_function) pour gérer vos promesses.

- If your add-in uses a CPU-intensive algorithm to process a large volume of data, you can use web workers to perform the long-running task in the background while running a separate script in the foreground, such as displaying progress in the user interface. Web workers do not block user activities and allow the HTML page to remain responsive. For an example of web workers, see [The Basics of Web Workers](https://www.html5rocks.com/tutorials/workers/basics/). See [Web Workers](https://developer.mozilla.org/docs/Web/API/Web_Workers_API) for more information about the Web Workers API.

- Si votre complément utilise un algorithme qui sollicite l’UC de manière intensive et si vous pouvez décomposer les entrées ou sorties de données en jeux de données de plus petite taille, créez un service web afin de lui passer les données et d’alléger la charge de l’UC, puis attendez un rappel asynchrone.

- Veillez à tester votre complément par rapport au volume de données le plus important possible, puis limitez votre complément pour lui permettre d’atteindre cette limite.

### <a name="performance-improvements-with-the-application-specific-apis"></a>Améliorations des performances avec les API spécifiques à l’application

Les conseils de performances relatifs à [l’utilisation du modèle d’API spécifique à l’application](../develop/application-specific-api-model.md) fournissent des conseils lors de l’utilisation des API spécifiques à l’application pour Excel, OneNote, Visio et Word. En résumé, vous devez :

- [Chargez uniquement les propriétés nécessaires](../develop/application-specific-api-model.md#calling-load-without-parameters-not-recommended).
- [Réduisez le nombre d’appels de synchronisation().](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-sync-calls) Lisez [Éviter d’utiliser la méthode context.sync dans les boucles](correlated-objects-pattern.md) pour plus d’informations sur la gestion `sync` des appels dans votre code.
- [Réduisez le nombre d’objets proxy créés](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-proxy-objects-created). Vous pouvez également annuler l’application des objets proxy, comme décrit dans la section suivante.

#### <a name="untrack-unneeded-proxy-objects"></a>Objets proxy inutiles non suivis

[Les objets proxy](../develop/application-specific-api-model.md#proxy-objects) sont conservés en mémoire jusqu’à ce qu’ils `RequestContext.sync()` sont appelés. Les opérations par lots volumineux peuvent générer un grand nombre d’objets proxy qui sont uniquement utiles une fois pour le complément et peuvent être publiés à partir de la mémoire avant l’exécution du lot.

La `untrack()` méthode libère l’objet de la mémoire. Cette méthode est implémentée sur de nombreux objets proxy d’API spécifiques à l’application. L’appel `untrack()` après que votre complément a été effectué avec l’objet doit générer un avantage notable en matière de performances lors de l’utilisation d’un grand nombre d’objets proxy.

> [!NOTE]
> `Range.untrack()` est un raccourci pour [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#office-officeextension-trackedobjects-remove-member(1)). N’importe quel objet proxy peut être non suivi en le supprimant de la liste d’objets suivis dans le contexte.

L’exemple de code Excel suivant remplit une plage sélectionnée avec des données, une cellule à la fois. Une fois que la valeur est ajoutée à la cellule, la plage représentant cette cellule est non suivie. Exécuter tout d’abord ce code avec une plage sélectionnée de 10 000 à 20 000 cellules, avec la `cell.untrack()` ligne et puis sans. Vous devez remarquer que le code est exécuté plus rapidement avec la `cell.untrack()` ligne que sans elle. Vous pouvez également remarquer un temps de réponse plus rapide par la suite, étant donné que l’étape de nettoyage prend moins de temps.

```js
Excel.run(async (context) => {
    const largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (let i = 0; i < largeRange.rowCount; i++) {
        for (let j = 0; j < largeRange.columnCount; j++) {
            let cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // Call untrack() to release the range from memory.
            cell.untrack();
        }
    }

    await context.sync();
});
```

Notez que le fait de devoir annuler l’utilisation d’objets devient important uniquement lorsque vous traitez avec des milliers d’entre eux. La plupart des compléments n’ont pas besoin de gérer le suivi des objets proxy.

## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Optimisation des performances à l’aide de l’API JavaScript d’Excel](../excel/performance.md)
