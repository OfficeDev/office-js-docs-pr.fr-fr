---
title: Limites des ressources et optimisation des performances pour les compléments Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: ead376bb12701f7ee810cfc4e536ae4866d2f1b5
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448139"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Limites des ressources et optimisation des performances pour les compléments Office

Afin d’offrir la meilleure expérience utilisateur, assurez-vous que votre complément Office fonctionne dans les limites prévues en matière d’utilisation du cœur du processeur et de la mémoire, ainsi qu’en matière de fiabilité et, pour les compléments Outlook, de temps de réponse lors de l’évaluation des expressions régulières. Ces limites propres à l’utilisation des ressources d’exécution s’appliquent aux compléments exécutés sur des clients Office pour Windows et OS X mais pas sur des clients Office Online, Outlook Web App ou OWA pour périphériques. 

Par ailleurs, cette rubrique suggère des techniques de conception et d’implémentation de complément qui permettent de mieux contrôler les performances des compléments sur ordinateurs de bureau et périphériques mobiles.

## <a name="resource-usage-limits-for-add-ins"></a>Limites d’utilisation des ressources pour les compléments

Les limites d’utilisation des ressources d’exécution s’appliquent à tous les compléments Office. Elles permettent à l’utilisateur de bénéficier de bonnes performances et contribuent également à atténuer les attaques par déni de service. Vous devez suffisamment tester votre complément Office sur l’application hôte cible à l’aide d’un large éventail de données possibles et mesurer ses performances par rapport aux limites ci-après :

- **Utilisation du cœur du processeur** - Seuil d’utilisation d’un seul cœur de processeur de 90 %, observé à trois reprises dans des intervalles par défaut de 5 secondes.

   L’intervalle par défaut de vérification de l’utilisation du cœur du processeur par un client enrichi de l’hôte est de 5 secondes. Si le client de l’hôte détecte que l’utilisation du cœur du processeur d’un complément dépasse la valeur du seuil, il affiche un message demandant à l’utilisateur s’il souhaite continuer à exécuter le complément. Si l’utilisateur choisit de continuer, le client de l’hôte ne redemande pas à l’utilisateur s’il souhaite continuer au cours de cette session de modification. Les administrateurs peuvent souhaiter utiliser la clé de registre **AlertInterval** pour augmenter le seuil et réduire l’affichage de ce message d’avertissement si les utilisateurs exécutent des compléments faisant appel au processeur de manière intensive.

- **Utilisation de la mémoire** - Seuil d’utilisation de mémoire par défaut, qui est déterminé de manière dynamique en fonction de la mémoire physique disponible de l’appareil.

   Par défaut, lorsqu’un client enrichi de l’hôte détecte que l’utilisation de la mémoire physique d’un appareil dépasse 80 % de la mémoire disponible, le client commence à surveiller l’utilisation de la mémoire du complément, au niveau du document pour les compléments du contenu et du volet des tâches, et au niveau de la boîte aux lettres pour les compléments Outlook. À un intervalle de 5 secondes par défaut, le client avertit l’utilisateur si l’utilisation de la mémoire physique pour un ensemble de compléments au niveau du document ou de la boîte aux lettres est supérieure à 50 %. Cette limite d’utilisation de la mémoire utilise la mémoire physique plutôt que la mémoire virtuelle pour garantir des performances sur des appareils dont la mémoire vive est limitée, comme les tablettes. Les administrateurs peuvent remplacer ce paramètre dynamique par une limite explicite en utilisant la clé de registre Windows **MemoryAlertThreshold** comme paramètre global, ou ajuster l’intervalle d’alerte en utilisant la clé **AlertInterval** comme paramètre global.

- **Tolérance d’incident** - Limite par défaut de 4 incidents pour un complément.

   Les administrateurs peuvent ajuster le seuil relatif aux incidents en utilisant la clé de registre **RestartManagerRetryLimit**.

- **Blocage d’application** - Limitation à 5 secondes du seuil de blocage prolongé d’un complément.

   Cette option affecte l’expérience utilisateur relative au complément et à l’application hôte. Dans ce cas, l’application hôte redémarre automatiquement tous les compléments actifs d’un document ou d’une boîte aux lettres (le cas échéant), et indique à l’utilisateur le complément qui ne répond pas. Les compléments peuvent atteindre ce seuil lorsqu’ils ne cèdent pas régulièrement le traitement lors de l’exécution de tâches longues. Il existe des techniques permettant de garantir qu’aucun blocage ne se produira. Les administrateurs ne peuvent pas remplacer ce seuil.

### <a name="outlook-add-ins"></a>Compléments Outlook

§LTA Si un complément Outlook dépasse les seuils précédents en matière d’utilisation du cœur du processeur ou de la mémoire, ou en matière de tolérance d’incident, Outlook désactive le complément. Le Centre d’administration Exchange indique que l’état de l’application est désactivé.

> [!NOTE]
> Même si seuls les clients enrichis Outlook et non les clients Outlook Web App ou OWA pour périphériques contrôlent l’utilisation des ressources, si un client enrichi désactive un complément Outlook, ce complément est également désactivé pour une utilisation dans Outlook Web App et OWA pour périphériques.

En plus du cœur du processeur, de la mémoire et des règles de fiabilité, les compléments Outlook doivent respecter les règles suivantes lors de l’activation :

- **Temps de réponse des expressions régulières** - Seuil par défaut de 1 000 millisecondes pour Outlook afin d’évaluer toutes les expressions régulières contenues dans le manifeste d’un complément Outlook. Le dépassement du seuil oblige Outlook à retenter l’évaluation un peu plus tard.

    À l’aide d’une stratégie de groupe ou d’un paramètre propre à l’application dans le registre Windows, les administrateurs peuvent ajuster cette valeur seuil par défaut de 1 000 millisecondes dans le paramètre **OutlookActivationAlertThreshold**.

- **Réévaluation des expressions régulières** - Limite par défaut de trois tentatives pour permettre à Outlook de réévaluer toutes les expressions régulières contenues dans un manifeste. Si l’évaluation échoue à trois reprises en dépassant le seuil applicable (qui est soit la valeur par défaut de 1 000 millisecondes, soit une valeur spécifiée par  **OutlookActivationAlertThreshold**, si ce paramètre existe dans le Registre Windows), Outlook désactive le complément Outlook. Le Centre d’administration Exchange affiche l’état désactivé. Par ailleurs, l’utilisation du complément est désactivée dans les clients riches Outlook, Outlook Web App et OWA pour les appareils.

    À l’aide d’une stratégie de groupe ou d’un paramètre propre à l’application dans le registre Windows, les administrateurs peuvent ajuster ce nombre de tentatives d’évaluation dans le paramètre **OutlookActivationManagerRetryLimit**.

### <a name="task-pane-and-content-add-ins"></a>Compléments de volet Office et de contenu

Si un complément de contenu ou de volet de tâches dépasse les seuils précédents en matière d’utilisation du cœur du processeur ou de la mémoire, ou en matière de tolérance d’incident, l’application hôte correspondante affiche un avertissement pour l’utilisateur. À ce stade, l’utilisateur peut effectuer l’une des actions suivantes :

- Redémarrer le complément.
- Annuler les alertes supplémentaires de dépassement de seuil. Dans l’idéal, l’utilisateur devrait supprimer le complément du document. La poursuite de l’exécution du complément risquerait d’entraîner des problèmes supplémentaires au niveau des performances et de la stabilité.  

## <a name="verifying-resource-usage-issues-in-the-telemetry-log"></a>Vérification des problèmes d’utilisation des ressources dans le journal de télémétrie

Office propose un journal de télémétrie qui tient à jour un enregistrement de certains événements (chargement, ouverture, fermeture et erreurs) des solutions Office qui s’exécutent sur l’ordinateur local, notamment les problèmes d’utilisation des ressources dans une Complément Office. Si le journal de télémétrie est configuré, vous pouvez utiliser Excel pour l’ouvrir à partir de l’emplacement par défaut suivant sur votre disque local :

`%Users%\<Current user>\AppData\Local\Microsoft\Office\15.0\Telemetry`

Le journal de télémétrie comprend pour chaque événement suivi pour un complément, les date/heure de l’occurrence, l’ID d’événement, la gravité et une courte description de l’événement, le nom convivial et l’ID unique du complément, ainsi que l’application qui a enregistré l’événement. Vous pouvez actualiser le journal de télémétrie pour afficher les événements suivis. Le tableau suivant répertorie des exemples de compléments Outlook qui ont été suivis dans le journal de télémétrie. 

|**Date/Heure**|**ID d’évènement**|**Gravité**|**Titre**|**Fichier**|**ID**|**Application**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|08/10/2012 17:57:10|7||Le manifeste du complément a été correctement téléchargé|Who’s Who|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|08/10/2012 17:57:01|7||Le manifeste du complément a été correctement téléchargé|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|

Le tableau suivant répertorie les événements que le journal de télémétrie suit pour les Compléments Office en général.

|**ID d’évènement**|**Titre**|**Gravité**|**Description**|
|:-----|:-----|:-----|:-----|
|7|Le manifeste du complément a été correctement téléchargé||Le manifeste de l’Complément Office a été chargé et lu correctement par l’application hôte.|
|8|Échec du téléchargement du manifeste du complément|Critique|L’application hôte n’a pas pu charger le fichier manifeste pour le complément Office à partir du catalogue SharePoint, du catalogue d’entreprise ou d’AppSource.|
|9|Impossible d’analyser le balisage du complément|Critique|L’application hôte a chargé le manifeste de l’Complément Office, mais n’a pas pu lire le balisage HTML de l’application.|
|10|Le complément a trop sollicité le processeur|Critique|L’Complément Office a utilisé plus de 90 % des ressources du processeur sur une période de temps définie.|
|15|Le complément a été désactivé en raison de l’expiration de la recherche de chaîne||Les compléments Outlook recherchent la ligne d’objet et le corps du message d’un courrier électronique pour déterminer s’ils doivent être affichés avec une expression régulière. Le complément Outlook répertorié dans la colonne  **Fichier** a été désactivé par Outlook, car il a expiré à plusieurs reprises lors d’une tentative de mise en correspondance d’une expression régulière.|
|18|Le complément a été fermé||L’application hôte a pu fermer l’Complément Office sans problème.|
|19|Le complément a rencontré une erreur d’exécution|Critique|L’Complément Office a rencontré un problème qui l’a empêchée de s’exécuter. Pour plus de détails, consultez le journal  **Alertes Microsoft Office** à l’aide de l’Observateur d’événements Windows sur l’ordinateur sur lequel l’erreur s’est produite.|
|20|Le complément n’a pas pu vérifier la licence|Critique|Les informations de licence de l’Complément Office n’ont pas pu être vérifiées et la licence a peut-être expiré. Pour plus de détails, consultez le journal  **Alertes Microsoft Office** à l’aide de l’Observateur d’événements Windows sur l’ordinateur sur lequel l’erreur s’est produite.|

Pour plus d’informations, consultez [Déployer le Tableau de bord de télémétrie](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15)) et [Dépannage des fichiers et des solutions personnalisées d’Office avec le journal de télémétrie](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log)


## <a name="design-and-implementation-techniques"></a>Techniques de conception et d’implémentation

Bien que les limites en matière d’utilisation des ressources de l’UC et de la mémoire, de tolérance d’incident et de réactivité de l’interface utilisateur s’appliquent aux Compléments Office qui s’exécutent uniquement sur les clients enrichis, l’optimisation de l’utilisation de ces ressources et de la batterie doit constituer une priorité si vous voulez que votre complément s’exécute de manière satisfaisante sur tous les clients et appareils de prise en charge. L’optimisation est particulièrement importante si votre complément effectue des opérations de longue durée ou manipule de grands jeux de données. La liste suivante suggère quelques techniques à suivre pour réduire la taille des opérations qui utilisent beaucoup de ressources d’UC ou de données afin de permettre à votre complément d’éviter une consommation excessive des ressources et à l’application hôte de rester réactive :

- Dans un scénario où votre complément a besoin de lire un important volume de données à partir d’un jeu de données illimité, vous pouvez appliquer la pagination lors de la lecture des données dans une table ou réduire la taille des données à chaque opération de lecture raccourcie, plutôt que de tenter de terminer la lecture en une seule opération. 

   For a JavaScript and jQuery code sample that shows breaking up a potentially long-running and CPU-intensive series of inputting and outputting operations on unbounded data, see [How can I give control back (briefly) to the browser during intensive JavaScript processing?](https://stackoverflow.com/questions/210821/how-can-i-give-control-back-briefly-to-the-browser-during-intensive-javascript). This example uses the [setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) method of the global object to limit the duration of input and output. It also handles the data in defined chunks instead of randomly unbounded data.

- Si votre complément utilise un algorithme qui sollicite l’UC de manière intensive pour traiter un important volume de données, vous pouvez recourir aux API Web Worker afin d’effectuer une tâche de longue durée en arrière-plan pendant qu’un script distinct s’exécute au premier plan (par exemple, l’affichage de la progression d’une opération dans l’interface utilisateur). Les API Web Worker ne bloquent pas les activités des utilisateurs. En outre, elles permettent à la page HTML de rester réactive. Pour obtenir un exemple d’API Web Worker, voir les [bases des API Web Worker](https://www.html5rocks.com/en/tutorials/workers/basics/). Pour plus d’informations sur l’API Web Worker Internet Explorer, voir [API Web Worker](https://developer.mozilla.org/docs/Web/API/Web_Workers_API).

- Si votre complément utilise un algorithme qui sollicite l’UC de manière intensive et si vous pouvez décomposer les entrées ou sorties de données en jeux de données de plus petite taille, créez un service web afin de lui passer les données et d’alléger la charge de l’UC, puis attendez un rappel asynchrone.

- Veillez à tester votre complément par rapport au volume de données le plus important possible, puis limitez votre complément pour lui permettre d’atteindre cette limite.


## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](/outlook/add-ins/limits-for-activation-and-javascript-api-for-outlook-add-ins)
- [Optimisation des performances à l’aide de l’API JavaScript d’Excel](../excel/performance.md)
