---
title: Pratiques recommandées et règles pour l’API de dialogue Office
description: Fournit des règles et des bonnes pratiques pour l’API de boîte de dialogue Office, telles que les meilleures pratiques pour une application monopage (SPA).
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: bdb92ba89faa63a5ca869be869f0a03cce91dba2
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958677"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Pratiques recommandées et règles pour l’API de dialogue Office

Cet article fournit des règles, des trousses et des bonnes pratiques pour l’API de boîte de dialogue Office, y compris les meilleures pratiques pour concevoir l’interface utilisateur d’une boîte de dialogue et utiliser l’API avec une application monopage (SPA)

> [!NOTE]
> Cet article suppose que vous êtes familiarisé avec les principes de base de l’utilisation de l’API de boîte de dialogue Office, comme décrit dans [Utiliser l’API de boîte de dialogue Office dans vos compléments Office](dialog-api-in-office-add-ins.md).
> 
> Voir aussi [Gestion des erreurs et des événements avec la boîte de dialogue Office](dialog-handle-errors-events.md).

## <a name="rules-and-gotchas"></a>Règles et pièges

- La boîte de dialogue peut uniquement accéder aux URL HTTPS, et non HTTP.
- L’URL transmise à la méthode [displayDialogAsync](/javascript/api/office/office.ui) doit se trouver dans le même domaine que le complément lui-même. Il ne peut pas s’agir d’un sous-domaine. Mais la page qui lui est passée peut rediriger vers une page d’un autre domaine.
- Une page hôte ne peut avoir qu’une seule boîte de dialogue ouverte à la fois. La page hôte peut être un volet Office ou le fichier de [fonction](/javascript/api/manifest/functionfile) d’une [commande de fonction](../design/add-in-commands.md#types-of-add-in-commands).
- Seules deux API Office peuvent être appelées dans la boîte de dialogue :
  - Fonction [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) .
  - `Office.context.requirements.isSetSupported` (Pour plus d’informations, consultez [Spécifier les applications Office et les exigences de l’API](specify-office-hosts-and-api-requirements.md).)
- La fonction [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) doit généralement être appelée à partir d’une page dans le même domaine que le complément lui-même, mais ce n’est pas obligatoire. Pour plus d’informations, consultez [Messagerie inter-domaines au runtime hôte](dialog-api-in-office-add-ins.md#cross-domain-messaging-to-the-host-runtime).

## <a name="best-practices"></a>Meilleures pratiques

### <a name="avoid-overusing-dialog-boxes"></a>Éviter la surutilisation des boîtes de dialogue

Comme des éléments d’interface utilisateur qui se chevauchent peuvent gêner des utilisateurs, évitez d’ouvrir une boîte de dialogue à partir d’un volet Office à moins que votre scénario l’exige. Lorsque vous envisagez d’utiliser la surface d’exposition d’un volet Office, tenez compte du fait que les volets Office peuvent être affichés sous forme d’onglets. Pour obtenir un exemple de volet office à onglets, consultez l’exemple [JavaScript SalesTracker du complément Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) .

### <a name="design-a-dialog-box-ui"></a>Concevoir une interface utilisateur de boîte de dialogue

Pour connaître les meilleures pratiques en matière de conception de boîte de dialogue, consultez [boîtes de dialogue dans les compléments Office](../develop/dialog-api-in-office-add-ins.md).

### <a name="handle-pop-up-blockers-with-office-on-the-web"></a>Gérer les bloqueurs contextuels avec Office sur le Web

Toute tentative d’affichage d’une boîte de dialogue lors de l’utilisation de Office sur le Web peut entraîner le blocage de la boîte de dialogue par le bloqueur de fenêtres contextuelles du navigateur. Si cela se produit, Office sur le Web ouvre une invite similaire à ce qui suit.

![Capture d’écran montrant l’invite avec une brève description et les boutons Autoriser et ignorer qu’un complément peut générer pour éviter les bloqueurs de fenêtres contextuelles dans le navigateur](../images/dialog-prompt-before-open.png)

Si l’utilisateur choisit **Autoriser**, la boîte de dialogue Office s’ouvre. Si l’utilisateur choisit **Ignorer**, l’invite se ferme et la boîte de dialogue Office ne s’ouvre pas. Au lieu de cela, la méthode retourne l’erreur `displayDialogAsync` 12009. Votre code doit intercepter cette erreur et fournir une autre expérience qui ne nécessite pas de boîte de dialogue, ou afficher un message à l’utilisateur indiquant que le complément exige qu’il autorise la boîte de dialogue. (Pour plus d’informations sur 12009, consultez [Erreurs de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).)

Si, pour une raison quelconque, vous souhaitez désactiver cette fonctionnalité, votre code doit refuser. Il effectue cette requête avec l’objet [DialogOptions](/javascript/api/office/office.dialogoptions) passé à la `displayDialogAsync` méthode. Plus précisément, l’objet doit inclure `promptBeforeOpen: false`. Lorsque cette option est définie sur false, Office sur le Web n’invite pas l’utilisateur à autoriser le complément à ouvrir une boîte de dialogue, et la boîte de dialogue Office ne s’ouvre pas.

### <a name="do-not-use-the-_host_info-value"></a>N’utilisez pas la valeur d’informations de l’hôte \_\_

Office ajoute automatiquement un paramètre de requête appelé `_host_info` à l’URL qui est transmise à `displayDialogAsync`. Il est ajouté après vos paramètres de requête personnalisés, le cas échéant. Elle n’est pas ajoutée aux URL suivantes auxquelles la boîte de dialogue accède. Microsoft peut modifier le contenu de cette valeur ou la supprimer entièrement, de sorte que votre code ne doit pas le lire. La même valeur est ajoutée au stockage de session de la boîte de dialogue (autrement dit, la propriété [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ). Là encore, *votre code ne doit ni lire, ni écrire cette valeur*.

### <a name="open-another-dialog-immediately-after-closing-one"></a>Ouvrir une autre boîte de dialogue immédiatement après la fermeture

Vous ne pouvez pas avoir plusieurs dialogues ouverts à partir d’une page hôte donnée. Votre code doit donc appeler [Dialog.close](/javascript/api/office/office.dialog#office-office-dialog-close-member(1)) sur une boîte de dialogue ouverte avant d’appeler `displayDialogAsync` pour ouvrir une autre boîte de dialogue. La `close` méthode est asynchrone. Pour cette raison, si vous appelez `displayDialogAsync` immédiatement après un appel de , la première boîte de `close`dialogue peut ne pas s’être complètement fermée quand Office tente d’ouvrir la seconde. Si cela se produit, Office renvoie une erreur [12007](dialog-handle-errors-events.md#12007) : « L’opération a échoué, car ce complément a déjà une boîte de dialogue active. »

La `close` méthode n’accepte pas de paramètre de rappel et ne retourne pas d’objet Promise. Elle ne peut donc pas être attendue avec le `await` mot clé ou avec une `then` méthode. Pour cette raison, nous vous suggérons la technique suivante lorsque vous devez ouvrir un nouveau dialogue immédiatement après la fermeture d’un dialogue : encapsulez le code pour ouvrir le nouveau dialogue dans une fonction et concevez la fonction pour qu’elle s’appelle de manière récursive si l’appel de `displayDialogAsync` retours `12007`. Voici un exemple.

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        openSecondDialog();
      }
      else {
         // Handle errors
      }
    }
  );
}
 
function openSecondDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
    (result) => {
      if(result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          openSecondDialog(); // Recursive call
        }
        else {
         // Handle other errors
        }
      }
    }
  );
}
```

Vous pouvez également forcer la suspension du code avant qu’il ne tente d’ouvrir le deuxième dialogue à l’aide de la méthode [setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) . Voici un exemple.

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        setTimeout(() => { 
          Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
             (result) => { /* callback body */ }
          );
        }, 1000);
      }
      else {
         // Handle errors
      }
    }
  );
}
```

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>Meilleures pratiques pour l’utilisation de l’API de boîte de dialogue Office dans une application spa

Si votre complément utilise le routage côté client, comme le font généralement les applications monopages, vous avez la possibilité de passer l’URL d’un itinéraire à la méthode [displayDialogAsync](/javascript/api/office/office.ui) au lieu de l’URL d’une page HTML distincte. *Nous vous déconseillons de le faire pour les raisons indiquées ci-dessous.*

> [!NOTE]
> Cet article n’est pas pertinent pour le routage *côté serveur* , par exemple dans une application web basée sur Express.

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>Problèmes liés aux contrats de fournisseur de services et à l’API de boîte de dialogue Office

La boîte de dialogue Office se trouve dans une nouvelle fenêtre avec sa propre instance du moteur JavaScript et, par conséquent, son propre contexte d’exécution complet. Si vous passez un itinéraire, votre page de base et tout son code d’initialisation et d’amorçage s’exécutent à nouveau dans ce nouveau contexte, et toutes les variables sont définies sur leurs valeurs initiales dans la boîte de dialogue. Par conséquent, cette technique télécharge et lance une deuxième instance de votre application dans la fenêtre box, ce qui va en partie à l’authentification monopage. En outre, le code qui modifie les variables dans la fenêtre de boîte de dialogue ne modifie pas la version du volet Office des mêmes variables. De même, la fenêtre de boîte de dialogue possède son propre stockage de session (propriété [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ), qui n’est pas accessible à partir du code dans le volet Office. La boîte de dialogue et la page hôte sur laquelle `displayDialogAsync` a été appelée ressemblent deux clients différents à votre serveur. (Pour un rappel de ce qu’est une page hôte, consultez [Ouvrir une boîte de dialogue à partir d’une page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).)

Ainsi, si vous avez passé un itinéraire à la `displayDialogAsync` méthode, vous n’auriez pas vraiment une SPA ; vous auriez *deux instances de la même SPA*. En outre, une grande partie du code de l’instance du volet Office n’est jamais utilisée dans cette instance et une grande partie du code de l’instance de boîte de dialogue ne sera jamais utilisée dans cette instance. Ce serait comme avoir deux SPAs dans le même lot.

#### <a name="microsoft-recommendations"></a>Recommandations Microsoft

Au lieu de passer un itinéraire côté client à la `displayDialogAsync` méthode, nous vous recommandons d’effectuer l’une des opérations suivantes :

* Si le code que vous souhaitez exécuter dans la boîte de dialogue est suffisamment complexe, créez explicitement deux SLA différentes ; autrement dit, avoir deux SPN dans des dossiers différents du même domaine. Une application spa s’exécute dans la boîte de dialogue et l’autre dans la page hôte de la boîte de dialogue où `displayDialogAsync` elle a été appelée. 
* Dans la plupart des scénarios, seule une logique simple est nécessaire dans la boîte de dialogue. Dans ce cas, votre projet sera considérablement simplifié en hébergeant une seule page HTML, avec JavaScript incorporé ou référencé, dans le domaine de votre spa. Passez l’URL de la page à la méthode`displayDialogAsync`. Bien que cela signifie que vous vous écartez de l’idée littérale d’une application monopage ; vous n’avez pas vraiment une seule instance d’une spa lorsque vous utilisez l’API de boîte de dialogue Office.
