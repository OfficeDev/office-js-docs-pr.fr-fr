---
title: Pratiques recommandées et règles pour l’API de dialogue Office
description: Fournit des règles et des pratiques recommandées pour l’API de boîte de dialogue Office, telles que les meilleures pratiques pour une application à page unique (SPA)
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: ffd609175276dc648805469847288fd2ff4f825c
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131786"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Pratiques recommandées et règles pour l’API de dialogue Office

Cet article fournit des règles, des pièges et des meilleures pratiques pour l’API de boîte de dialogue Office, notamment les meilleures pratiques pour la conception de l’interface utilisateur d’une boîte de dialogue et l’utilisation de l’API avec dans une application à page unique (SPA)

> [!NOTE]
> Cet article suppose que vous êtes familiarisé avec les notions de base de l’utilisation de l’API de boîte de dialogue Office, comme décrit dans la rubrique [use the Office Dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).
> 
> Voir aussi [gestion des erreurs et des événements à l’aide de la boîte de dialogue Office](dialog-handle-errors-events.md).

## <a name="rules-and-gotchas"></a>Règles et pièges

- La boîte de dialogue ne peut accéder qu’aux URL HTTPs, et non à HTTP.
- L’URL transmise à la méthode [displayDialogAsync](/javascript/api/office/office.ui) doit se trouver dans le même domaine que le complément lui-même. Il ne peut pas s’agir d’un sous-domaine. Mais la page qui lui est transmise peut rediriger vers une page dans un autre domaine.
- Une fenêtre hôte, qui peut être un volet de tâches ou le fichier de [fonction](../reference/manifest/functionfile.md) sans interface utilisateur d’une commande de complément, ne peut avoir qu’une seule boîte de dialogue ouverte à la fois.
- Seules deux API Office peuvent être appelées dans la boîte de dialogue :
  - La fonction [messageParent](/javascript/api/office/office.ui#messageparent-message-) .
  - `Office.context.requirements.isSetSupported` (Pour plus d’informations, consultez la rubrique [spécifier les applications Office et les conditions requises](specify-office-hosts-and-api-requirements.md)de l’API.)
- La fonction [messageParent](/javascript/api/office/office.ui#messageparent-message-) peut uniquement être appelée à partir d’une page dans le même domaine que le complément lui-même.

## <a name="best-practices"></a>Meilleures pratiques

### <a name="avoid-overusing-dialog-boxes"></a>Éviter de surutiliser les boîtes de dialogue

Comme des éléments d’interface utilisateur qui se chevauchent peuvent gêner des utilisateurs, évitez d’ouvrir une boîte de dialogue à partir d’un volet Office à moins que votre scénario l’exige. Lorsque vous envisagez d’utiliser la surface d’exposition d’un volet Office, tenez compte du fait que les volets Office peuvent être affichés sous forme d’onglets. Pour voir un exemple, consultez la rubrique relative à l’exemple [Complément Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

### <a name="designing-a-dialog-box-ui"></a>Conception d’une interface utilisateur de boîte de dialogue

Pour obtenir les meilleures pratiques en matière de conception de boîte de dialogue, consultez la rubrique [boîtes de dialogue dans les compléments Office](../design/dialog-boxes.md).

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a>Gestion des bloqueurs de fenêtres publicitaires avec Office sur le web

La tentative d’affichage d’une boîte de dialogue lors de l’utilisation d’Office sur le Web peut entraîner le blocage de la boîte de dialogue par le bloqueur de fenêtres publicitaires du navigateur. Office sur le Web est doté d’une fonctionnalité qui permet aux boîtes de dialogue de votre complément d’être une exception au bloqueur de fenêtres publicitaires intempestives du navigateur. Lorsque votre code appelle la `displayDialogAsync` méthode, Office sur le Web ouvre une invite semblable à la suivante.

![Capture d’écran illustrant l’invite avec une brève description et les boutons autoriser et ignorer qu’un complément peut générer pour éviter les bloqueurs de fenêtres publicitaires intempestives dans le navigateur](../images/dialog-prompt-before-open.png)

Si l’utilisateur choisit **autoriser**, la boîte de dialogue Office s’ouvre. Si l’utilisateur choisit **Ignorer**, l’invite se ferme et la boîte de dialogue Office ne s’ouvre pas. Au lieu de cela, la `displayDialogAsync` méthode renvoie l’erreur 12009. Votre code doit intercepter cette erreur et fournir une autre expérience qui ne nécessite pas de boîte de dialogue ou afficher un message à l’utilisateur pour lui demander d’autoriser la boîte de dialogue. (Pour plus d’informations sur 12009, voir [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).)

Si, pour une raison quelconque, vous souhaitez désactiver cette fonctionnalité, votre code doit l’exclure. Il effectue cette demande avec l’objet [DialogOptions](/javascript/api/office/office.dialogoptions) qui est transmis à la `displayDialogAsync` méthode. Plus précisément, l’objet doit inclure `promptBeforeOpen: false` . Lorsque cette option est définie sur false, Office sur le Web n’invite pas l’utilisateur à autoriser le complément à ouvrir une boîte de dialogue et la boîte de dialogue Office ne s’ouvre pas.

### <a name="do-not-use-the-_host_info-value"></a>Ne pas utiliser la \_ \_ valeur Host info

Office ajoute automatiquement un paramètre de requête appelé `_host_info` à l’URL qui est transmise à `displayDialogAsync`. Il est ajouté après vos paramètres de requête personnalisés, le cas échéant. Il n’est ajouté à aucune URL suivante vers laquelle la boîte de dialogue navigue. Microsoft peut modifier le contenu de cette valeur ou le supprimer entièrement, de sorte que votre code ne doit pas le lire. La même valeur est ajoutée à l’espace de stockage de session de la boîte de dialogue (autrement dit, la propriété [Window. sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ). Là encore, *votre code ne doit ni lire, ni écrire cette valeur*.

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>Meilleures pratiques pour l’utilisation de l’API de boîte de dialogue Office dans un SPA

Si votre complément utilise le routage côté client, comme le font généralement des applications à page unique (SPAs), vous avez la possibilité de transmettre l’URL d’un itinéraire à la méthode [displayDialogAsync](/javascript/api/office/office.ui) au lieu de l’URL d’une page HTML distincte. *Nous vous recommandons de le faire pour les raisons indiquées ci-dessous.*

> [!NOTE]
> Cet article n’est pas pertinent pour le routage *côté serveur* , comme dans une application Web basée sur Express.

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>Problèmes liés à la fonction SPAs et l’API de boîte de dialogue Office

La boîte de dialogue Office se trouve dans une nouvelle fenêtre avec sa propre instance du moteur JavaScript, ce qui lui est propre. Si vous transmettez un itinéraire, votre page de base et toutes ses initialisations et codes d’amorçage s’exécutent à nouveau dans ce nouveau contexte, et toutes les variables sont définies sur leurs valeurs initiales dans la boîte de dialogue. Ainsi, cette technique télécharge et lance une deuxième instance de votre application dans la fenêtre de zone, ce qui annule partiellement l’objectif d’un SPA. De plus, le code qui modifie les variables dans la fenêtre de la boîte de dialogue ne modifie pas la version du volet Office des mêmes variables. De même, la fenêtre de la boîte de dialogue dispose de son propre espace de stockage de session (propriété [Window. sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ), qui n’est pas accessible à partir du code dans le volet Office. La boîte de dialogue et la page hôte sur laquelle l' `displayDialogAsync` appel a été appelé ressemblent à deux clients différents sur votre serveur. (Pour un rappel de ce qu’est une page hôte, voir [ouvrir une boîte de dialogue à partir d’une page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).)

Par conséquent, si vous avez passé un itinéraire à la `displayDialogAsync` méthode, vous n’auriez pas vraiment un spa ; vous auriez *deux instances du même Spa*. De plus, la majeure partie du code dans l’instance de volet de tâches ne serait jamais utilisée dans cette instance et la plus grande partie du code dans l’instance de boîte de dialogue ne seraient jamais utilisées dans cette instance. Ce serait comme avoir deux SPAs dans le même lot.

#### <a name="microsoft-recommendations"></a>Recommandations Microsoft

Au lieu de transmettre un itinéraire côté client à la `displayDialogAsync` méthode, nous vous recommandons d’effectuer l’une des opérations suivantes :

* Si le code que vous souhaitez exécuter dans la boîte de dialogue est suffisamment complexe, créez deux autres SPAs de manière explicite ; autrement dit, avoir deux SPAs dans différents dossiers du même domaine. Un SPA s’exécute dans la boîte de dialogue et l’autre dans la page hôte de la boîte de dialogue où `displayDialogAsync` a été appelé. 
* Dans la plupart des scénarios, seule une logique simple est nécessaire dans la boîte de dialogue. Dans ce cas, votre projet est grandement simplifié en hébergeant une page HTML unique, avec JavaScript incorporé ou référencé, dans le domaine de votre SPA. Passez l’URL de la page à la méthode`displayDialogAsync`. Cela signifie que vous vous écartez de l’idée littérale d’une application à page unique ; vous n’avez pas réellement une seule instance d’un SPA lorsque vous utilisez l’API de boîte de dialogue Office.
