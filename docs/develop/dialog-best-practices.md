---
title: Pratiques recommandées et règles pour l’API de dialogue Office
description: 'Fournit des règles et des meilleures pratiques pour Office API de boîte de dialogue, telles que les meilleures pratiques pour une application mono-page (SPA)'
ms.date: 07/22/2021
ms.localizationpriority: medium
---

# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Pratiques recommandées et règles pour l’API de dialogue Office

Cet article fournit des règles, des gotchas et des meilleures pratiques pour l’API de boîte de dialogue Office, y compris les meilleures pratiques pour la conception de l’interface utilisateur d’une boîte de dialogue et l’utilisation de l’API dans une application mono-page (SPA)

> [!NOTE]
> Cet article présuppose que vous connaissez les principes de base de l’utilisation de l’API de boîte de dialogue Office, comme décrit dans Utiliser [l’API](dialog-api-in-office-add-ins.md) de boîte de dialogue Office dans vos Office.
> 
> Voir aussi [Gestion des erreurs et des événements à l’Office boîte de dialogue.](dialog-handle-errors-events.md)

## <a name="rules-and-gotchas"></a>Règles et pièges

- La boîte de dialogue peut uniquement accéder aux URL HTTPS, et non à HTTP.
- L’URL transmise à la [méthode displayDialogAsync](/javascript/api/office/office.ui) doit se trouver exactement dans le même domaine que le add-in lui-même. Il ne peut pas s’agit d’un sous-domaine. Toutefois, la page qui lui est transmise peut rediriger vers une page d’un autre domaine.
- Une fenêtre hôte, qui peut être un volet Des tâches ou le fichier de fonction sans interface [](../reference/manifest/functionfile.md) utilisateur d’une commande de add-in, ne peut avoir qu’une seule boîte de dialogue ouverte à la fois.
- Seules deux Office API peuvent être appelées dans la boîte de dialogue :
  - Fonction [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) .
  - `Office.context.requirements.isSetSupported`(Pour plus d’informations, voir [Spécifier les Office applications et les api requises](specify-office-hosts-and-api-requirements.md).)
- La [fonction messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) doit généralement être appelée à partir d’une page dans le même domaine que le module lui-même, mais cela n’est pas obligatoire. Pour plus d’informations, consultez [Messagerie inter-domaines au runtime hôte](dialog-api-in-office-add-ins.md#cross-domain-messaging-to-the-host-runtime).

## <a name="best-practices"></a>Meilleures pratiques

### <a name="avoid-overusing-dialog-boxes"></a>Éviter toute surutilisation des boîtes de dialogue

Comme des éléments d’interface utilisateur qui se chevauchent peuvent gêner des utilisateurs, évitez d’ouvrir une boîte de dialogue à partir d’un volet Office à moins que votre scénario l’exige. Lorsque vous envisagez d’utiliser la surface d’exposition d’un volet Office, tenez compte du fait que les volets Office peuvent être affichés sous forme d’onglets. Pour obtenir un exemple de volet De tâches à onglets, voir [l’exemple Excel javaScript SalesTracker du add-in](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

### <a name="design-a-dialog-box-ui"></a>Concevoir une interface utilisateur de boîte de dialogue

Pour obtenir les meilleures pratiques en matière de conception de boîte de dialogue, voir Boîtes de dialogue [Office des applications](../design/dialog-boxes.md).

### <a name="handle-pop-up-blockers-with-office-on-the-web"></a>Gérer les bloqueurs de fenêtres Office sur le Web

Toute tentative d’affichage d’une boîte de dialogue lors de l’Office sur le Web peut entraîner le blocage de la boîte de dialogue par le bloqueur de fenêtres d’affichage du navigateur. Si cela se produit, Office sur le Web ouvre une invite semblable à celle-ci.

![Capture d’écran montrant l’invite avec une brève description et les boutons Autoriser et Ignorer qu’un add-in peut générer pour éviter les bloqueurs de fenêtres pop-up dans le navigateur](../images/dialog-prompt-before-open.png)

Si l’utilisateur choisit **Autoriser**, la boîte Office dialogue s’ouvre. Si l’utilisateur choisit **Ignorer**, l’invite se ferme et la boîte Office dialogue ne s’ouvre pas. Au lieu de cela, la `displayDialogAsync` méthode renvoie l’erreur 12009. Votre code doit capturer cette erreur et fournir une expérience de remplacement qui ne nécessite pas de boîte de dialogue ou afficher un message à l’utilisateur pour lui conseiller que le add-in exige qu’il autorise la boîte de dialogue. (Pour plus d’informations sur 12009, voir [Erreurs de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).)

Si, pour une raison quelconque, vous souhaitez désactiver cette fonctionnalité, votre code doit le désactiver. Il effectue cette demande avec [l’objet DialogOptions](/javascript/api/office/office.dialogoptions) qui est transmis à la `displayDialogAsync` méthode. Plus précisément, l’objet doit inclure `promptBeforeOpen: false`. Lorsque cette option est définie sur False, Office sur le Web demande pas à l’utilisateur d’autoriser le add-in à ouvrir une boîte de dialogue et la boîte de dialogue Office ne s’ouvre pas.

### <a name="do-not-use-the-_host_info-value"></a>N’utilisez pas la valeur \_hostinfo\_

Office ajoute automatiquement un paramètre de requête appelé `_host_info` à l’URL qui est transmise à `displayDialogAsync`. Il est appended après vos paramètres de requête personnalisés, le cas cas. Elle n’est pas appendée aux URL suivantes vers qui la boîte de dialogue navigue. Microsoft peut modifier le contenu de cette valeur ou le supprimer entièrement, de sorte que votre code ne doit pas le lire. La même valeur est ajoutée au stockage de session de la boîte de dialogue (autrement dit, la [propriété Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ). Là encore, *votre code ne doit ni lire, ni écrire cette valeur*.

### <a name="open-another-dialog-immediately-after-closing-one"></a>Ouvrir une autre boîte de dialogue immédiatement après en avoir fermé une

Comme plusieurs boîtes de dialogue ne peuvent pas être ouvertes à partir d’une page hôte donnée, votre code doit appeler [Dialog.close](/javascript/api/office/office.dialog#office-office-dialog-close-member(1)) `displayDialogAsync` dans une boîte de dialogue ouverte avant d’appeler l’ouverture d’une autre boîte de dialogue. La `close` méthode est asynchrone. Pour cette raison, `displayDialogAsync` `close`si vous appelez immédiatement après un appel de , il se peut que la première boîte de dialogue ne soit pas complètement fermée Office tente d’ouvrir la seconde. Si cela se produit, Office renvoyer une erreur [12007](dialog-handle-errors-events.md#12007) : « L’opération a échoué, car ce module a déjà une boîte de dialogue active. »

La `close` méthode n’accepte pas de paramètre de rappel et ne retourne pas d’objet Promise `await` , elle ne peut donc pas être attendue avec le mot clé ou avec une `then` méthode. Pour cette raison, nous vous suggérons la technique suivante lorsque vous devez ouvrir une nouvelle boîte de dialogue immédiatement après la fermeture d’une boîte de dialogue : encapsulez le code pour ouvrir la nouvelle boîte de dialogue dans une méthode et concevez la méthode pour qu’elle s’appelle de manière récursive `displayDialogAsync` `12007`si l’appel de retour . Voici un exemple.

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

Vous pouvez également forcer l’interruption du code avant d’essayer d’ouvrir la deuxième boîte de dialogue à l’aide de la [méthode setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) . Voici un exemple.

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

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>Meilleures pratiques pour l’utilisation de l’API Office boîte de dialogue dans une SPA

Si votre application utilise le routage côté client, comme le font généralement les applications mono-page, vous avez la possibilité de transmettre l’URL d’un itinéraire à la méthode [displayDialogAsync](/javascript/api/office/office.ui) au lieu de l’URL d’une page HTML distincte. *Nous vous déconseillons de le faire pour les raisons ci-dessous.*

> [!NOTE]
> Cet article n’est pas pertinent pour *le routage* côté serveur, comme dans une application web express.

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>Problèmes avec les spa et l’API Office boîte de dialogue de gestion

La Office boîte de dialogue se trouve dans une nouvelle fenêtre avec sa propre instance du moteur JavaScript, et par conséquent son propre contexte d’exécution complet. Si vous passez un itinéraire, votre page de base et tout son code d’initialisation et de mise en route s’exécutent à nouveau dans ce nouveau contexte, et toutes les variables sont définies sur leurs valeurs initiales dans la boîte de dialogue. Par conséquent, cette technique télécharge et lance une deuxième instance de votre application dans la fenêtre box, ce qui va partiellement à l’emploi d’une SPA. En outre, le code qui modifie des variables dans la fenêtre de boîte de dialogue ne modifie pas la version du volet Des tâches des mêmes variables. De même, la fenêtre de la boîte de dialogue possède son propre stockage de session (propriété [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ), qui n’est pas accessible à partir du code dans le volet Des tâches. La boîte de dialogue et la page hôte sur laquelle `displayDialogAsync` a été appelée ressemblent deux clients différents à votre serveur. (Pour un rappel de ce qu’est une page hôte, voir Ouvrir une boîte de [dialogue à partir d’une page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).)

Par exemple, si vous avez transmis un itinéraire `displayDialogAsync` à la méthode, vous n’ariez pas vraiment de SPA ; vous ariez deux *instances de la même SPA*. En outre, une grande partie du code dans l’instance du volet Des tâches ne sera jamais utilisée dans cette instance et la plus grande partie du code dans l’instance de la boîte de dialogue ne sera jamais utilisée dans cette instance. Ce serait comme avoir deux SPAs dans le même lot.

#### <a name="microsoft-recommendations"></a>Recommandations de Microsoft

Au lieu de transmettre un itinéraire côté client `displayDialogAsync` à la méthode, nous vous recommandons d’adopter l’une des méthodes suivantes :

* Si le code que vous souhaitez exécuter dans la boîte de dialogue est suffisamment complexe, créez explicitement deux spa différents . autrement dit, avoir deux spas dans des dossiers différents du même domaine. Une SPA s’exécute dans la boîte de dialogue et l’autre dans la page `displayDialogAsync` hôte de la boîte de dialogue où elle a été appelée. 
* Dans la plupart des scénarios, seule une logique simple est nécessaire dans la boîte de dialogue. Dans ce cas, votre projet sera considérablement simplifié en hébergeant une page HTML unique, avec javaScript incorporé ou référencé, dans le domaine de votre SPA. Passez l’URL de la page à la méthode`displayDialogAsync`. Cela signifie que vous déviez de l’idée littérale d’une application à page unique ; vous n’avez pas vraiment une seule instance d’une SPA lorsque vous utilisez l’API Office dialogue.
