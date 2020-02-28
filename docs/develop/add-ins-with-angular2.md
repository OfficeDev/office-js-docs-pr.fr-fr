---
title: Développement de compléments Office avec Angular
description: ''
ms.date: 01/27/2020
localization_priority: Normal
ms.openlocfilehash: 0e78a2e000016e28f2012902e0e8db487b80d930
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324637"
---
# <a name="develop-office-add-ins-with-angular"></a>Développement de compléments Office avec Angular

Cet article fournit des conseils sur l’utilisation d’Angular 2+ pour créer un complément Office sous la forme d’une application monopage.

> [!NOTE]
> Avez-vous une contribution à apporter suite à votre expérience d’utilisation d’Angular pour créer des compléments Office ? Vous pouvez contribuer à cet article dans [GitHub](https://github.com/OfficeDev/office-js-docs) ou fournir vos commentaires en envoyant un [problème](https://github.com/OfficeDev/office-js-docs-pr/issues) dans le référentiel. 

Pour obtenir un exemple de complément Office créé à l’aide de l’infrastructure Angular, consultez [Complément de vérification du style dans Word basé sur Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="install-the-typescript-type-definitions"></a>Installer les définitions de type TypeScript

Ouvrez une fenêtre nodejs et entrez les informations suivantes sur la ligne de commande :

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a>L’amorçage doit s’effectuer à l’intérieur d’Office.initialize

Sur n’importe quelle page qui appelle les API JavaScript Office, Word ou Excel, votre code doit tout d’abord affecter une `Office.initialize` méthode à la propriété. (Si vous n’avez pas de code d’initialisation, le corps de la méthode peut`{}`être vide «», mais vous ne devez `Office.initialize` pas laisser la propriété undefined. Pour plus d’informations, consultez [la rubrique initialiser votre complément Office](initialize-add-in.md). Office appelle cette méthode immédiatement après avoir initialisé les bibliothèques JavaScript Office.

**Votre code d’amorçage Angular doit être appelé à l’intérieur de la méthode que vous affectez à `Office.initialize`** pour vous assurer que les bibliothèques JavaScript Office ont été initialisées en premier. Voici un exemple simple qui montre comment procéder. Ce code doit figurer dans le fichier main.ts du projet.

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a>Utiliser la stratégie d’emplacement de hachage dans l’application Angular

La navigation entre des itinéraires dans l’application peut ne pas fonctionner si vous ne spécifiez pas la stratégie d’emplacement de hachage. Vous pouvez procéder de deux manières. Tout d’abord, vous pouvez spécifier un fournisseur pour la stratégie d’emplacement dans le module de votre application, comme montré dans l’exemple suivant. Il est placé dans le fichier app.module.ts.

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity

@NgModule({
  providers: [
    { provide: LocationStrategy, useClass: HashLocationStrategy },
    // Other providers suppressed
  ],
  // Other module properties suppressed
})
export class AppModule { }
``` 

Si vous définissez vos itinéraires dans un module de routage distinct, il existe une autre façon de spécifier la stratégie d’emplacement de hachage. Dans le fichier .ts de votre module de routage, passez un objet de configuration vers la fonction `forRoot` qui spécifie la stratégie. Voici un exemple de code. 

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity

const routes: Routes = // route definitions go here

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
```


## <a name="consider-wrapping-fabric-components-with-angular-components"></a>Insertion de composants Fabric dans des composants Angular

Nous vous recommandons d’utiliser le style [UI Fabric](https://developer.microsoft.com/fabric#) dans votre complément. L’interface utilisateur Fabric pour le Web est disponible en deux versions : 

- [Fabric React](https://developer.microsoft.com/fabric#/controls/web) fournit des composants fiables, à jour et qui sont extrêmement personnalisables.

- [Fabric Core](https://developer.microsoft.com/fabric#/styles/web) est un ensemble de classes CSS et de mixins Sass qui vous permettent d’accéder aux couleurs, animations, polices, icônes et grilles de Fabric.

Envisagez d’utiliser des composants de structure dans votre complément en les insérant dans les composants Angular. Pour obtenir un exemple de procédure à suivre, voir [Complément de vérification du style dans Word basé sur Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Notez, par exemple, comment le composant Angular défini dans [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) importe le fichier de structure TextField.ts, où le composant de structure est défini. 


## <a name="using-the-office-dialog-api-with-angular"></a>Utilisation de l’API de boîte de dialogue Office avec Angular

L’API de boîte de dialogue du complément Office permet à votre complément d’ouvrir une page dans une boîte de dialogue non modale dans laquelle vous pouvez échanger des informations avec la page principale, qui se trouve généralement dans un volet Office.

La méthode [displayDialogAsync](/javascript/api/office/office.ui) accepte un paramètre qui indique l’URL de la page qui doit s’ouvrir dans la boîte de dialogue. Votre complément peut avoir une autre page HTML (différente de la page de base) pour passer à ce paramètre, ou vous pouvez passer l’URL d’un itinéraire dans votre application Angular. 

Il est important de ne pas oublier, si vous passez un itinéraire, que la boîte de dialogue crée une nouvelle fenêtre avec son propre contexte d’exécution. Votre page de base et son code d’initialisation et d’amorçage s’exécutent à nouveau dans ce nouveau contexte, et toutes les variables sont définies sur leurs valeurs initiales dans la boîte de dialogue. Par conséquent, cette technique lance une deuxième instance de votre application monopage dans la boîte de dialogue. Le code qui modifie des variables dans la boîte de dialogue ne change pas la version du volet Office des mêmes variables. De même, la boîte de dialogue possède son propre stockage de session, qui n’est pas accessible à partir du code dans le volet Office.  


## <a name="trigger-the-ui-update"></a>Déclencher la mise à jour de l’interface utilisateur

Dans une application Angula, l’interface utilisateur ne se met parfois pas à jour. Cela est dû au fait que cette partie du code s’exécute en dehors de la zone Angular. La solution consiste à placer le code dans la zone, comme montré dans l’exemple suivant.

```js
import { NgZone } from '@angular/core';

export class MyComponent {
  constructor(private zone: NgZone) { }

  myFunction() {
    this.zone.run(() => {
      // the codes that need update the UI
    });
  }
}
```

## <a name="using-observable"></a>Utilisation d’un élément Observable

Angular utilise RxJS (Reactive Extensions for JavaScript), et RxJS présente les objets `Observable` et `Observer` pour implémenter le traitement asynchrone. Cette section fournit une brève introduction à l’utilisation de `Observables` ; pour plus d’informations, consultez la documentation [RxJS](https://rxjs-dev.firebaseapp.com/) officielle.

Un `Observable` est semblable à un objet `Promise` d’une certaine façon - il est renvoyé immédiatement à partir d’un appel asynchrone, mais il ne peut être résolu qu’après un certain délai. Toutefois, bien qu’une `Promise` soit une valeur unique (qui peut être un objet de tableau), un `Observable` est un tableau d’objets (éventuellement avec un seul membre). Cela permet d’appeler les [méthodes de tableaux](https://www.w3schools.com/jsref/jsref_obj_array.asp), telles que `concat`, `map` et `filter`, sur des objets `Observable`. 

### <a name="pushing-instead-of-pulling"></a>Poussée au lieu d’extraction

Votre code « pousse » les objets `Promise` en les affectant aux variables, mais les objets `Observable` « poussent » leurs valeurs vers les objets qui *s’abonnent* à l’objet `Observable`. Les abonnés sont des objets `Observer`. L’avantage de l’architecture Push est que les nouveaux membres peuvent être ajoutés au tableau `Observable` au fil du temps. Lorsqu’un nouveau membre est ajouté, tous les objets `Observer` qui s’abonnent à `Observable` reçoivent une notification. 

L’`Observer` est configuré pour traiter chaque nouvel objet (appelé l’objet « suivant ») avec une fonction. (Il est également configuré pour répondre à une erreur et à une notification d’achèvement. Consultez la section suivante pour obtenir un exemple.) Pour cette raison, les objets `Observable` peuvent être utilisés dans un plus large éventail de scénarios que les objets `Promise`. Par exemple, en plus de retourner un `Observable` à partir d’un appel AJAX, de la façon dont vous pouvez retourner une `Promise`, un `Observable` peut être renvoyé à partir d’un gestionnaire d’événements, tel que le gestionnaire d’événements « modifié » pour une zone de texte. Chaque fois qu’un utilisateur saisit du texte dans la zone, tous les objets `Observer` abonnés réagissent immédiatement en utilisant le dernier texte et/ou l’état actuel de l’application en tant qu’entrée. 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a>Attendre jusqu'à ce que tous les appels asynchrones soient terminés

Lorsque vous voulez vous assurer qu’un rappel ne s’exécute que lorsque tous les membres d’un ensemble d’objets `Promise` sont résolus, utilisez la méthode `Promise.all()`.

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
``` 

Pour faire la même chose avec un objet `Observable`, vous utilisez la méthode [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md).  

```js
const source = Observable.forkJoin([x, y, z]);

const subscription = source.subscribe(
  x => {
    // TODO: Callback logic goes here
  },
  err => console.log('Error: ' + err),
  () => console.log('Completed')
);
``` 

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a>Compiler l’application Angular à l’aide du compilateur Ahead-of-Time (AOT)

Les performances de l’application représentent l’un des aspects importants de l’expérience utilisateur. Une application Angular peut être optimisée à l’aide du compilateur Ahead-of-Time (AOT) d’Angular pour compiler l’application au moment de la génération. Le compilateur convertit tout le code source (modèles HTML et TypeScript) en code JavaScript efficace. Si vous compilez votre application avec le compilateur AOT, aucune autre compilation ne se produira pendant l’exécution. Ainsi, le rendu est plus rapide et les requêtes asynchrones sont plus rapides pour les modèles HTML. Par ailleurs, la taille globale de l’application sera réduite, car le compilateur d’Angular n’a pas besoin d’être inclus dans le distribuable de l’application. 

Pour utiliser le compilateur AOT, ajoutez `--aot` à la commande `ng build` ou `ng serve` :

```command&nbsp;line
ng build --aot
ng serve --aot
```

> [!NOTE]
> Pour en savoir plus sur le compilateur Ahead-of-Time (AOT) d’Angular, consultez le [guide officiel](https://angular.io/guide/aot-compiler).
