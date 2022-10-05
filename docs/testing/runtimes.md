---
title: Runtimes dans les compléments Office
description: Découvrez les runtimes utilisés par les compléments Office.
ms.date: 09/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: c20845e6df6adfa2fa382f10e8c7f5de786aeab8
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467229"
---
# <a name="runtimes-in-office-add-ins"></a>Runtimes dans les compléments Office

Les compléments Office s’exécutent dans des runtimes incorporés dans Office. En tant que langage interprété, JavaScript doit s’exécuter dans un runtime JavaScript. [Node.js](https://nodejs.org) et les navigateurs modernes sont des exemples de ces runtimes. 

## <a name="types-of-runtimes"></a>Types de runtimes

Il existe deux types de runtimes utilisés par les compléments Office :

- **Runtime JavaScript uniquement** : moteur JavaScript complété par la prise en charge de [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [cors complet (partage de ressources cross-origin)](https://developer.mozilla.org/docs/Web/HTTP/CORS) et stockage côté client des données. Il ne prend pas en charge le [stockage local](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou les cookies.
- **Runtime de navigateur** : inclut toutes les fonctionnalités d’un runtime JavaScript uniquement et ajoute la prise en charge du [stockage local](https://developer.mozilla.org/docs/Web/API/Window/localStorage), un [moteur de rendu](https://developer.mozilla.org/docs/Glossary/Rendering_engine) qui affiche du code HTML et des cookies.

Vous trouverez plus de détails sur ces types plus loin dans cet article sur le [runtime JavaScript uniquement](#javascript-only-runtime) et le [runtime browser](#browser-runtime).

Le tableau suivant indique les fonctionnalités possibles d’un complément qui utilisent chaque type d’exécution. 

| Type d’exécution | Fonctionnalité de complément |
|:-----|:-----|
| JavaScript uniquement | [Fonctions personnalisées](../excel/custom-functions-overview.md) Excel</br>(sauf lorsque le runtime est [partagé](#shared-runtime) ou que le complément est en cours d’exécution dans Office sur le Web)</br></br>[Tâche basée sur les événements Outlook](../outlook/autolaunch.md)</br>(uniquement lorsque le complément est en cours d’exécution dans Outlook sur Windows)|
| Navigateur | [volet Office](../design/task-pane-add-ins.md)</br></br>[fenêtre de dialogue](../develop/dialog-api-in-office-add-ins.md)</br></br>[commande de fonction](../design/add-in-commands.md#types-of-add-in-commands)</br></br>[Fonctions personnalisées](../excel/custom-functions-overview.md) Excel</br>(lorsque le runtime est [partagé](#shared-runtime) ou que le complément est en cours d’exécution dans Office sur le Web)</br></br>[Tâche basée sur les événements Outlook](../outlook/autolaunch.md)</br>(lorsque le complément s’exécute dans Outlook sur Mac ou Outlook sur le web)|

Le tableau suivant présente les mêmes informations organisées par le type de runtime utilisé pour les différentes fonctionnalités possibles d’un complément.

| Fonctionnalité de complément | Type d’exécution sur Windows | Type d’exécution sur Mac | Type de runtime sur le web |
|:-----|:-----|:-----|:-----|
|Fonctions personnalisées dans Excel | JavaScript uniquement</br>(mais *navigateur* lorsque le runtime est partagé)|JavaScript uniquement</br>(mais *navigateur* lorsque le runtime est partagé)| Navigateur |
|Tâches basées sur des événements Outlook | JavaScript uniquement | Navigateur | Navigateur |
|volet Office | Navigateur | Navigateur | Navigateur |
|fenêtre de dialogue | Navigateur | Navigateur | Navigateur |
|commande de fonction | Navigateur | Navigateur | Navigateur |


Dans Office sur le Web, tout s’exécute toujours dans un runtime de type navigateur. En fait, à une exception près, tout ce qui se trouve dans un complément sur le web s’exécute dans le *même* processus de navigateur : le processus de navigateur dans lequel l’utilisateur a ouvert Office sur le Web. L’exception est quand une boîte de dialogue est ouverte avec un appel [d’Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) et que l’option [DialogOptions.displayInIFrame](/javascript/api/office/office.dialogoptions#office-office-dialogoptions-displayiniframe-member) *n’est pas* passée et définie sur `true`. Lorsque l’option n’est pas passée (elle a donc la valeur par défaut `false` ), la boîte de dialogue s’ouvre dans son propre processus. Le même principe s’applique à la méthode [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) et à l’option [OfficeRuntime.DisplayWebDialogOptions.displayInIFrame](/javascript/api/office-runtime/officeruntime.displaywebdialogoptions#office-runtime-officeruntime-displaywebdialogoptions-displayiniframe-member) .

Lorsqu’un complément s’exécute sur une plateforme autre que le web, les principes suivants s’appliquent.

- Une boîte de dialogue s’exécute dans son propre processus d’exécution. 
- Une tâche basée sur les événements Outlook s’exécute dans son propre processus d’exécution. 
- Par défaut, les volets office, les commandes de fonction et les fonctions personnalisées Excel s’exécutent chacun dans leur propre processus d’exécution. Toutefois, pour certaines applications hôtes Office, le manifeste de complément peut être configuré afin que deux, ou les trois, puissent s’exécuter dans le même runtime. Consultez [le runtime partagé](#shared-runtime).

Selon l’application Office hôte et les fonctionnalités utilisées dans le complément, il peut y avoir de nombreux runtimes dans un complément. Chacun s’exécute généralement dans son propre processus, mais pas nécessairement simultanément. Les éléments suivants sont des exemples.

- Un complément PowerPoint ou Word qui ne partage aucun runtime et inclut les fonctionnalités suivantes a jusqu’à trois runtimes.

  - Volet Office
  - Commande de fonction
  - Boîte de dialogue (une boîte de dialogue peut être lancée à partir du volet Office ou de la commande de fonction.) 
  
      > [!NOTE]
      > Il n’est pas recommandé d’ouvrir plusieurs dialogues simultanément, mais si le complément permet à l’utilisateur d’en ouvrir une à partir du volet Office et une autre à partir de la commande de fonction en même temps, ce complément aurait quatre runtimes. Un volet Office et un appel donné d’une commande de fonction ne peuvent avoir qu’une seule boîte de dialogue ouverte à la fois ; mais si la commande de fonction est appelée plusieurs fois, une nouvelle boîte de dialogue s’ouvre au-dessus de son prédécesseur à chaque appel, de sorte qu’il peut y avoir de nombreux runtimes. Le reste de cette liste ignore la possibilité de plusieurs dialogues ouverts.

- Un complément Excel qui ne partage aucun runtime et inclut les fonctionnalités suivantes a *jusqu’à quatre* runtimes.

  - Volet Office
  - Commande de fonction
  - Une fonction personnalisée
  - Boîte de dialogue (une boîte de dialogue peut être lancée à partir du volet Office, de la commande de fonction ou d’une fonction personnalisée.)

- Un complément Excel avec les mêmes fonctionnalités et est configuré pour partager le même runtime dans le volet Office, la commande de fonction et la fonction personnalisée, a *deux* runtimes. Un runtime partagé ne peut ouvrir qu’une seule boîte de dialogue à la fois.
- Un complément Excel avec les mêmes fonctionnalités, sauf qu’il n’a pas de boîte de dialogue et qu’il est configuré pour partager le même runtime dans le volet Office, la commande de fonction et la fonction personnalisée, a *un runtime* .
- Un complément Outlook qui comporte les fonctionnalités suivantes a *jusqu’à quatre* runtimes. (Les runtimes ne peuvent pas être partagés dans Outlook.)

  - Volet Office
  - Commande de fonction
  - Une tâche basée sur des événements
  - Boîte de dialogue (une boîte de dialogue peut être lancée à partir du volet Office ou de la commande de fonction, mais pas à partir d’une tâche basée sur un événement.)

## <a name="share-data-across-runtimes"></a>Partager des données entre les runtimes

> [!NOTE]
> - Si vous savez que votre complément sera utilisé uniquement dans Office sur le Web et qu’il n’ouvre aucune boîte de dialogue avec l’option `displayInIFrame` définie sur `true`, vous pouvez ignorer cette section. Étant donné que tous les éléments de votre complément s’exécutent dans le même processus d’exécution, vous pouvez simplement utiliser des variables globales pour partager des données entre les fonctionnalités.
> - Comme indiqué ci-dessus dans [Types de runtimes](#types-of-runtimes), le type d’exécution utilisé par une fonctionnalité varie en partie selon la plateforme. Il est recommandé d’éviter d’avoir du code de complément basé sur la plateforme. Par conséquent, les conseils de cette section recommandent des techniques qui fonctionnent sur plusieurs plateformes. Il n’existe qu’un seul cas, indiqué ci-dessous, dans lequel le code de branchement est requis. 

Pour les compléments Excel, PowerPoint et Word, utilisez un [runtime partagé](#shared-runtime) lorsque deux fonctionnalités ou plus, à l’exception des boîtes de dialogue, doivent partager des données. Dans Outlook, ou dans les scénarios où le partage d’un runtime n’est pas possible, vous avez besoin d’autres méthodes. Les parties du complément qui se trouvent dans des processus d’exécution distincts ne partagent pas automatiquement les données globales et sont traitées par le serveur d’applications web du complément comme des sessions distinctes. [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ne peut donc pas être utilisé pour partager des données entre eux. *Les conseils suivants supposent que vous n’utilisez pas un runtime partagé.*

- Transmettez des données entre une boîte de dialogue et son volet office parent, la commande de fonction ou la fonction personnalisée à l’aide des méthodes [Office.ui.messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) et [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) . 

    > [!NOTE]
    > Les `OfficeRuntime.storage` méthodes ne peuvent pas être appelées dans une boîte de dialogue. Il ne s’agit donc pas d’une option permettant de partager des données entre un dialogue et un autre runtime. 

- Pour partager des données entre un volet Office et une commande de fonction, stockez les données dans [Window.localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), qui est partagé entre tous les runtimes qui accèdent à la même [origine](https://developer.mozilla.org/docs/Glossary/Origin) spécifique. 
    > [!NOTE]
    > LocalStorage n’est pas accessible dans un runtime JavaScript uniquement et, par conséquent, il n’est pas disponible dans les fonctions personnalisées Excel. Il ne peut pas non plus être utilisé pour partager des données avec des tâches basées sur des événements Outlook (car ces tâches utilisent un runtime JavaScript uniquement sur certaines plateformes).

    > [!TIP]
    > Les données in `Window.localStorage` persistent entre les sessions du complément et sont partagées par des compléments de la même origine. Ces deux caractéristiques sont souvent indésirables pour un complément. 
    >
    > - Pour vous assurer que chaque session d’un complément donné démarre un nouvel appel de la méthode [Window.localStorage.clear](https://developer.mozilla.org/docs/Web/API/Storage/clear) au démarrage du complément. 
    > - Pour conserver certaines valeurs stockées, mais réinitialiser d’autres valeurs, utilisez [Window.localStorage.setItem](https://developer.mozilla.org/docs/Web/API/Storage/setItem) lorsque le complément démarre pour chaque élément qui doit être réinitialisé à une valeur initiale. 
    > - Pour supprimer entièrement un élément, appelez [Window.localStorage.removeItem](https://developer.mozilla.org/docs/Web/API/Storage/removeItem).

- Pour partager des données entre une fonction personnalisée Excel et tout autre runtime, utilisez [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage).
- Pour partager des données entre une tâche basée sur des événements Outlook et une commande de volet office ou de fonction, vous devez brancher votre code en fonction de la valeur de la propriété [Office.context.platform](/javascript/api/office/office.context#office-office-context-platform-member) . 

    - Lorsque la valeur est `PC` (Windows), stockez et récupérez des données à l’aide des API [Office.sessionData](/javascript/api/outlook/office.sessiondata) .
    - Lorsque la valeur est `Mac`, utilisez-la `Window.localStorage` comme décrit précédemment dans cette liste.

Voici d’autres façons de partager des données :

- Stockez les données partagées dans une base de données en ligne accessible à tous les runtimes.
- Stockez les données partagées dans un cookie pour le domaine du complément afin de les partager entre les runtimes du navigateur. Les runtimes JavaScript uniquement ne prennent pas en charge les cookies.

Pour plus d’informations, consultez [Conserver l’état et les paramètres du complément](../develop/persisting-add-in-state-and-settings.md) , ainsi que [Gérer l’état et les paramètres d’un complément Outlook](../outlook/manage-state-and-settings-outlook.md).

## <a name="javascript-only-runtime"></a>Runtime JavaScript uniquement

Le runtime JavaScript uniquement utilisé dans les compléments Office est une modification d’un runtime open source créé à l’origine pour [React Native](https://reactnative.dev/). Il contient un moteur JavaScript complété par la prise en charge de [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [CORS complet (partage de ressources cross-origin)](https://developer.mozilla.org/docs/Web/HTTP/CORS) et [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage). Il n’a pas de moteur de rendu et ne prend pas en charge les cookies ni le [stockage local](https://developer.mozilla.org/docs/Web/API/Window/localStorage). 

Ce type de runtime est utilisé dans les tâches basées sur les événements Outlook dans Office sur Windows uniquement et dans les fonctions personnalisées Excel *, sauf* lorsque les fonctions personnalisées [partagent un runtime](#shared-runtime). 

- Lorsqu’il est utilisé pour une fonction personnalisée Excel, le runtime démarre lorsque la feuille de calcul est recalculé ou que la fonction personnalisée calcule. Il ne s’arrête pas tant que le classeur n’est pas fermé.  
- Lorsqu’il est utilisé dans une tâche basée sur des événements Outlook, le runtime démarre lorsque l’événement se produit. Elle se termine lorsque la première des opérations suivantes se produit.

  - Le gestionnaire d’événements appelle la `completed` méthode de son paramètre d’événement.
  - 5 minutes se sont écoulées depuis l’événement déclencheur.
  - L’utilisateur modifie le focus à partir de la fenêtre où l’événement a été déclenché, par exemple une fenêtre de composition de message.

Un runtime JavaScript utilise moins de mémoire et démarre plus rapidement qu’un runtime de navigateur, mais offre moins de fonctionnalités.

## <a name="browser-runtime"></a>Runtime du navigateur

Les compléments Office utilisent un autre runtime de type de navigateur en fonction de la plateforme dans laquelle Office s’exécute (web, Mac ou Windows), ainsi que de la version et de la build de Windows et Office. Par exemple, si l’utilisateur exécute Office sur le Web dans un navigateur FireFox, le runtime Firefox est utilisé. Si l’utilisateur exécute Office sur Mac, le runtime Safari est utilisé. Si l’utilisateur exécute Office sur Windows, un Edge ou Internet Explorer fournit le runtime, en fonction de la version de Windows et d’Office. Vous trouverez plus d’informations dans [les navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

Tous ces runtimes incluent un moteur de rendu HTML et prennent en charge [webSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [CORS complet (partage de ressources cross-origin)](https://developer.mozilla.org/docs/Web/HTTP/CORS) et [le stockage local](https://developer.mozilla.org/docs/Web/API/Window/localStorage) et les cookies.

La durée de vie d’un runtime de navigateur varie en fonction de la fonctionnalité qu’il implémente et de son partage ou non.

- Lorsqu’un complément avec un volet Office est lancé, un runtime de navigateur démarre, sauf s’il s’agit d’un runtime partagé qui est déjà en cours d’exécution. S’il s’agit d’un runtime partagé, il s’arrête lorsque le document est fermé. S’il ne s’agit pas d’un runtime partagé, il s’arrête lorsque le volet Office est fermé.
- Lorsqu’une boîte de dialogue est ouverte, un runtime de navigateur démarre. Elle s’arrête lorsque la boîte de dialogue est fermée.
- Lorsqu’une commande de fonction est exécutée (ce qui se produit lorsqu’un utilisateur sélectionne son bouton ou son élément de menu), un runtime de navigateur démarre, sauf s’il s’agit d’un runtime partagé qui est déjà en cours d’exécution. S’il s’agit d’un runtime partagé, il s’arrête lorsque le document est fermé. S’il ne s’agit pas d’un runtime partagé, il s’arrête lorsque le premier des événements suivants se produit.
 
  - La commande de fonction appelle la `completed` méthode de son paramètre d’événement.
  - 5 minutes se sont écoulées depuis l’événement déclencheur. (Si une boîte de dialogue a été ouverte dans la commande de fonction et qu’elle est toujours ouverte lorsque le runtime parent expire, le runtime de dialogue reste en cours d’exécution jusqu’à ce que la boîte de dialogue soit fermée.)

- Lorsqu’une fonction personnalisée Excel utilise un runtime partagé, un runtime de type navigateur démarre lorsque la fonction personnalisée calcule si le runtime partagé n’a pas encore démarré pour une autre raison. Il s’arrête lorsque le document est fermé.

> [!NOTE]
> Lorsqu’un runtime est [partagé](#shared-runtime), il est possible pour votre code de fermer le volet Office sans arrêter le complément. Pour plus d’informations, consultez [afficher ou masquer le volet Office de votre complément Office](../develop/show-hide-add-in.md) .

Un runtime de navigateur a plus de fonctionnalités qu’un runtime JavaScript uniquement, mais démarre plus lentement et utilise plus de mémoire.

### <a name="shared-runtime"></a>Runtime partagé requis

Un « runtime partagé » n’est pas un type de runtime. Il fait référence à un [runtime de type navigateur](#browser-runtime) qui est partagé par les fonctionnalités du complément qui auraient autrement leur propre runtime. Plus précisément, vous avez la possibilité de configurer le volet Office et les commandes de fonction du complément pour partager un runtime. Dans un complément Excel, vous pouvez également configurer des fonctions personnalisées pour partager le runtime d’un volet office ou d’une commande de fonction, ou les deux. Dans ce cas, les fonctions personnalisées s’exécutent dans un runtime de type navigateur, au lieu d’un [runtime JavaScript uniquement](#javascript-only-runtime) comme dans le cas contraire. Consultez [Configurer votre complément pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md) pour plus d’informations sur les avantages et les limitations du partage des runtimes et des instructions pour configurer le complément de manière à utiliser un runtime partagé. En bref, le runtime JavaScript uniquement utilise moins de mémoire et démarre plus rapidement, mais dispose de moins de fonctionnalités.

> [!NOTE]
> - Vous pouvez partager des runtimes uniquement dans Excel, PowerPoint et Word. 
> - Vous ne pouvez pas configurer une boîte de dialogue pour partager un runtime. Chaque boîte de dialogue a toujours sa propre, sauf lorsque le dialogue est lancé dans Office sur le Web avec l’option `displayInIFrame` définie sur `true`.
> - Un runtime partagé n’utilise jamais le runtime Microsoft Edge WebView (EdgeHTML) d’origine. Si les conditions d’utilisation de Microsoft Edge avec WebView2 (basée sur Chromium) sont remplies (comme spécifié dans [les navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md)), ce runtime est utilisé. Sinon, le runtime Internet Explorer 11 est utilisé.