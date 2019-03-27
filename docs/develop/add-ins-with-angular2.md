---
title: Développement de compléments Office avec Angular
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 20e66fe1a7a6cce17428424290d4e4c99995efb2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872227"
---
# <a name="develop-office-add-ins-with-angular"></a><span data-ttu-id="6dcf3-102">Développement de compléments Office avec Angular</span><span class="sxs-lookup"><span data-stu-id="6dcf3-102">Develop Office Add-ins with Angular</span></span>

<span data-ttu-id="6dcf3-103">Cet article fournit des conseils sur l’utilisation d’Angular 2+ pour créer un complément Office sous la forme d’une application monopage.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-103">This article provides guidance for using Angular 2+ to create an Office Add-in as a single page application.</span></span>

> [!NOTE]
> <span data-ttu-id="6dcf3-p101">Avez-vous une contribution à apporter suite à votre expérience d’utilisation d’Angular pour créer des compléments Office ? Vous pouvez contribuer à cet article dans [GitHub](https://github.com/OfficeDev/office-js-docs) ou fournir vos commentaires en envoyant un [problème](https://github.com/OfficeDev/office-js-docs-pr/issues) dans le référentiel.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p101">Do you have something to contribute based on your experience using Angular to create Office Add-ins? You can contribute to this article in [GitHub](https://github.com/OfficeDev/office-js-docs) or provide your feedback by submitting an [issue](https://github.com/OfficeDev/office-js-docs-pr/issues) in the repo.</span></span> 

<span data-ttu-id="6dcf3-106">Pour obtenir un exemple de complément Office créé à l’aide de l’infrastructure Angular, consultez [Complément de vérification du style dans Word basé sur Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).</span><span class="sxs-lookup"><span data-stu-id="6dcf3-106">For an Office Add-ins sample that's built using the Angular framework, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).</span></span>

## <a name="install-the-typescript-type-definitions"></a><span data-ttu-id="6dcf3-107">Installer les définitions de type TypeScript</span><span class="sxs-lookup"><span data-stu-id="6dcf3-107">Install the TypeScript type definitions</span></span>

<span data-ttu-id="6dcf3-108">Ouvrez une fenêtre nodejs et entrez les informations suivantes sur la ligne de commande :</span><span class="sxs-lookup"><span data-stu-id="6dcf3-108">Open an nodejs window and enter the following at the command line:</span></span>

```bash
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a><span data-ttu-id="6dcf3-109">L’amorçage doit s’effectuer à l’intérieur d’Office.initialize</span><span class="sxs-lookup"><span data-stu-id="6dcf3-109">Bootstrapping must be inside Office.initialize</span></span>

<span data-ttu-id="6dcf3-p102">Dans une page qui appelle les API Office, Word ou Excel JavaScript, votre code doit d’abord attribuer une méthode à la propriété `Office.initialize`. (Si vous ne possédez aucun code d’initialisation, le corps de la méthode peut contenir simplement des symboles « `{}` » vides, mais vous ne devez pas laisser la propriété `Office.initialize` non définie. Pour plus d’informations, voir [Initialisation de votre complément](understanding-the-javascript-api-for-office.md#initializing-your-add-in).) Office appelle cette méthode immédiatement après l’initialisation des bibliothèques JavaScript Office.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p102">On any page that calls the Office, Word, or Excel JavaScript APIs, your code must first assign a method to the `Office.initialize` property. (If you have no initialization code, the method body can be just empty "`{}`" symbols, but you must not leave the `Office.initialize` property undefined. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).) Office calls this method immediately after it has initialized the Office JavaScript libraries.</span></span>

<span data-ttu-id="6dcf3-p103">**Votre code d’amorçage Angular doit être appelé à l’intérieur de la méthode que vous affectez à `Office.initialize`** pour vous assurer que les bibliothèques JavaScript Office ont été initialisées en premier. Voici un exemple simple qui montre comment procéder. Ce code doit figurer dans le fichier main.ts du projet.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p103">**Your Angular bootstrapping code must be called inside the method that you assign to `Office.initialize`** to ensure that the Office JavaScript libraries have initialized first. The following is a simple example that shows how to do this. This code should be in the main.ts file of the project.</span></span>

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a><span data-ttu-id="6dcf3-116">Utiliser la stratégie d’emplacement de hachage dans l’application Angular</span><span class="sxs-lookup"><span data-stu-id="6dcf3-116">Use the hash location strategy in the Angular application</span></span>

<span data-ttu-id="6dcf3-p104">La navigation entre des itinéraires dans l’application peut ne pas fonctionner si vous ne spécifiez pas la stratégie d’emplacement de hachage. Vous pouvez procéder de deux manières. Tout d’abord, vous pouvez spécifier un fournisseur pour la stratégie d’emplacement dans le module de votre application, comme montré dans l’exemple suivant. Il est placé dans le fichier app.module.ts.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p104">Navigating between routes in the application might not work if you don't specify the hash location strategy. You can do this in one of two ways. First, you can specify a provider for the location strategy in your app module, as shown in the following example. It goes into the app.module.ts file.</span></span>

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

<span data-ttu-id="6dcf3-p105">Si vous définissez vos itinéraires dans un module de routage distinct, il existe une autre façon de spécifier la stratégie d’emplacement de hachage. Dans le fichier .ts de votre module de routage, passez un objet de configuration vers la fonction `forRoot` qui spécifie la stratégie. Voici un exemple de code.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p105">If you define your routes in a separate routing module, there is an alternative way to specify the hash location strategy. In your routing module's .ts file, pass a configuration object to the `forRoot` function that specifies the strategy. The following code is an example.</span></span> 

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


## <a name="consider-wrapping-fabric-components-with-angular-components"></a><span data-ttu-id="6dcf3-124">Insertion de composants Fabric dans des composants Angular</span><span class="sxs-lookup"><span data-stu-id="6dcf3-124">Consider wrapping Fabric components with Angular components</span></span>

<span data-ttu-id="6dcf3-p106">Nous vous recommandons d’utiliser le style de la [structure de l’interface utilisateur d’Office](https://developer.microsoft.com/fabric#/fabric-js) dans votre complément. La structure inclut des composants qui sont fournis dans différentes versions, notamment une version [basée sur TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Envisagez d’utiliser des composants de structure dans votre complément en les insérant dans les composants Angular. Pour obtenir un exemple de procédure à suivre, voir [Complément de vérification du style dans Word basé sur Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Notez, par exemple, comment le composant Angular défini dans [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) importe le fichier de structure TextField.ts, où le composant de structure est défini.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p106">We recommend using [Office UI Fabric](https://developer.microsoft.com/fabric#/fabric-js) styling in your add-in. Fabric includes components that come in several versions, including a version [based on TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Consider using Fabric components in your add-in by wrapping them in Angular components. For an example that shows you how to do this, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Note, for example, how the Angular component defined in [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) imports the Fabric file TextField.ts, where the Fabric component is defined.</span></span> 


## <a name="using-the-office-dialog-api-with-angular"></a><span data-ttu-id="6dcf3-130">Utilisation de l’API Boîte de dialogue Office</span><span class="sxs-lookup"><span data-stu-id="6dcf3-130">Using the Office Dialog API with Angular</span></span>

<span data-ttu-id="6dcf3-131">L’API Boîte de dialogue du complément Office permet à votre complément d’ouvrir une page dans une boîte de dialogue semi-modale dans laquelle vous pouvez échanger des informations avec la page principale, qui se trouve généralement dans un volet Office.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-131">The Office Add-in Dialog API enables your add-in to open a page in a semimodal dialog box that can exchange information with the main page, which is typically in a task pane.</span></span>

<span data-ttu-id="6dcf3-p107">La méthode [displayDialogAsync](/javascript/api/office/office.ui) accepte un paramètre qui indique l’URL de la page qui doit s’ouvrir dans la boîte de dialogue. Votre complément peut avoir une autre page HTML (différente de la page de base) pour passer à ce paramètre, ou vous pouvez passer l’URL d’un itinéraire dans votre application Angular.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p107">The [displayDialogAsync](/javascript/api/office/office.ui) method takes a parameter that specifies the URL of the page that should open in the dialog box. Your add-in can have a separate HTML page (different from the base page) to pass to this parameter, or you can pass the URL of a route in your Angular appication.</span></span> 

<span data-ttu-id="6dcf3-p108">Il est important de ne pas oublier, si vous passez un itinéraire, que la boîte de dialogue crée une nouvelle fenêtre avec son propre contexte d’exécution. Votre page de base et son code d’initialisation et d’amorçage s’exécutent à nouveau dans ce nouveau contexte, et toutes les variables sont définies sur leurs valeurs initiales dans la boîte de dialogue. Par conséquent, cette technique lance une deuxième instance de votre application monopage dans la boîte de dialogue. Le code qui modifie des variables dans la boîte de dialogue ne change pas la version du volet Office des mêmes variables. De même, la boîte de dialogue possède son propre stockage de session, qui n’est pas accessible à partir du code dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p108">It is important to remember, if you pass a route, that the dialog box creates a new window with its own execution context. Your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog box. So this technique launches a second instance of your single page application in the dialog box. Code that changes variables in the dialog box does not change the task pane version of the same variables. Similarly, the dialog box has its own session storage, which is not accessible from code in the task pane.</span></span>  


## <a name="trigger-the-ui-update"></a><span data-ttu-id="6dcf3-139">Déclencher la mise à jour de l’interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="6dcf3-139">Trigger the UI update</span></span>

<span data-ttu-id="6dcf3-p109">Dans une application Angula, l’interface utilisateur ne se met parfois pas à jour. Cela est dû au fait que cette partie du code s’exécute en dehors de la zone Angular. La solution consiste à placer le code dans la zone, comme montré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p109">In an Angular app, the UI sometimes does not update. This is because that part of the code runs out of the Angular zone. The solution is to put the code in the zone, as shown in the following example.</span></span>

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

## <a name="using-observable"></a><span data-ttu-id="6dcf3-143">Utilisation d’un élément Observable</span><span class="sxs-lookup"><span data-stu-id="6dcf3-143">Using Observable</span></span>

<span data-ttu-id="6dcf3-p110">Angular utilise RxJS (Reactive Extensions for JavaScript), et RxJS présente les objets `Observable` et `Observer` pour implémenter le traitement asynchrone. Cette section fournit une brève introduction à l’utilisation de `Observables` ; pour plus d’informations, consultez la documentation [RxJS](https://rxjs-dev.firebaseapp.com/) officielle.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p110">Angular uses RxJS (Reactive Extensions for JavaScript), and RxJS introduces `Observable` and `Observer` objects to implement asynchronous processing. This section provides a brief introduction to using `Observables`; for more detailed information, see the official [RxJS](https://rxjs-dev.firebaseapp.com/) documentation.</span></span>

<span data-ttu-id="6dcf3-p111">Un `Observable` est semblable à un objet `Promise` d’une certaine façon - il est renvoyé immédiatement à partir d’un appel asynchrone, mais il ne peut être résolu qu’après un certain délai. Toutefois, bien qu’une `Promise` soit une valeur unique (qui peut être un objet de tableau), un `Observable` est un tableau d’objets (éventuellement avec un seul membre). Cela permet d’appeler les [méthodes de tableaux](https://www.w3schools.com/jsref/jsref_obj_array.asp), telles que `concat`, `map` et `filter`, sur des objets `Observable`.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p111">An `Observable` is like a `Promise` object in some ways - it is returned immediately from an asynchronous call, but it might not resolve until some time later. However, while a `Promise` is a single value (which can be an array object), an `Observable` is an array of objects (possibly with only a single member). This enables code to call [array methods](https://www.w3schools.com/jsref/jsref_obj_array.asp), such as `concat`, `map`, and `filter`, on `Observable` objects.</span></span> 

### <a name="pushing-instead-of-pulling"></a><span data-ttu-id="6dcf3-149">Poussée au lieu d’extraction</span><span class="sxs-lookup"><span data-stu-id="6dcf3-149">Pushing instead of pulling</span></span>

<span data-ttu-id="6dcf3-p112">Votre code « pousse » les objets `Promise` en les affectant aux variables, mais les objets `Observable` « poussent » leurs valeurs vers les objets qui *s’abonnent* à l’objet `Observable`. Les abonnés sont des objets `Observer`. L’avantage de l’architecture Push est que les nouveaux membres peuvent être ajoutés au tableau `Observable` au fil du temps. Lorsqu’un nouveau membre est ajouté, tous les objets `Observer` qui s’abonnent à `Observable` reçoivent une notification.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p112">Your code "pulls" `Promise` objects by assigning them to variables, but `Observable` objects "push" their values to objects that *subscribe* to the `Observable`. The subscribers are `Observer` objects. The benefit of the push architecture is that new members can be added to the `Observable` array over time. When a new member is added, all the `Observer` objects that subscribe to the `Observable` receive a notification.</span></span> 

<span data-ttu-id="6dcf3-p113">L’`Observer` est configuré pour traiter chaque nouvel objet (appelé l’objet « suivant ») avec une fonction. (Il est également configuré pour répondre à une erreur et à une notification d’achèvement. Consultez la section suivante pour obtenir un exemple.) Pour cette raison, les objets `Observable` peuvent être utilisés dans un plus large éventail de scénarios que les objets `Promise`. Par exemple, en plus de retourner un `Observable` à partir d’un appel AJAX, de la façon dont vous pouvez retourner une `Promise`, un `Observable` peut être renvoyé à partir d’un gestionnaire d’événements, tel que le gestionnaire d’événements « modifié » pour une zone de texte. Chaque fois qu’un utilisateur saisit du texte dans la zone, tous les objets `Observer` abonnés réagissent immédiatement en utilisant le dernier texte et/ou l’état actuel de l’application en tant qu’entrée.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-p113">The `Observer` is configured to process each new object (called the "next" object) with a function. (It is also configured to respond to an error and a completion notification. See the next section for an example.) For this reason, `Observable` objects can be used in a wider range of scenarios than `Promise` objects. For example, in addition to returning an `Observable` from an AJAX call, the way you can return a `Promise`, an `Observable` can be returned from an event handler, such as the "changed" event handler for a text box. Each time a user enters text in the box, all the subscribed `Observer` objects react immediately using the latest text and/or the current state of the application as input.</span></span> 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a><span data-ttu-id="6dcf3-159">Attendre jusqu'à ce que tous les appels asynchrones soient terminés</span><span class="sxs-lookup"><span data-stu-id="6dcf3-159">Waiting until all asynchronous calls have completed</span></span>

<span data-ttu-id="6dcf3-160">Lorsque vous voulez vous assurer qu’un rappel ne s’exécute que lorsque tous les membres d’un ensemble d’objets `Promise` sont résolus, utilisez la méthode `Promise.all()`.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-160">When you want to ensure that a callback only runs when every member of a set of `Promise` objects has resolved, use the `Promise.all()` method.</span></span>

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
``` 

<span data-ttu-id="6dcf3-161">Pour faire la même chose avec un objet `Observable`, vous utilisez la méthode [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md).</span><span class="sxs-lookup"><span data-stu-id="6dcf3-161">To do the same thing with an `Observable` object, you use the [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) method.</span></span>  

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

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a><span data-ttu-id="6dcf3-162">Compiler l’application Angular à l’aide du compilateur Ahead-of-Time (AOT)</span><span class="sxs-lookup"><span data-stu-id="6dcf3-162">Compile the Angular application using the Ahead-of-Time (AOT) compiler</span></span>

<span data-ttu-id="6dcf3-163">Les performances de l’application représentent l’un des aspects importants de l’expérience utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-163">Application performance is one of the most important aspects of user experience.</span></span> <span data-ttu-id="6dcf3-164">Une application Angular peut être optimisée à l’aide du compilateur Ahead-of-Time (AOT) d’Angular pour compiler l’application au moment de la génération.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-164">An Angular application can be optimized by using the Angular Ahead-of-Time (AOT) compiler to compile the app at build time.</span></span> <span data-ttu-id="6dcf3-165">Le compilateur convertit tout le code source (modèles HTML et TypeScript) en code JavaScript efficace.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-165">It converts all source code (HTML templates and TypeScript) into efficient JavaScript code.</span></span> <span data-ttu-id="6dcf3-166">Si vous compilez votre application avec le compilateur AOT, aucune autre compilation ne se produira pendant l’exécution. Ainsi, le rendu est plus rapide et les requêtes asynchrones sont plus rapides pour les modèles HTML.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-166">If you compile your app with the AOT compiler, no additional compilation will occur at runtime, which results in faster rendering and faster asynchronous requests for HTML templates.</span></span> <span data-ttu-id="6dcf3-167">Par ailleurs, la taille globale de l’application sera réduite, car le compilateur d’Angular n’a pas besoin d’être inclus dans le distribuable de l’application.</span><span class="sxs-lookup"><span data-stu-id="6dcf3-167">Additionally, the overall application size will be reduced, because the Angular compiler won't need to be included in the application distributable.</span></span> 

<span data-ttu-id="6dcf3-168">Pour utiliser le compilateur AOT, ajoutez `--aot` à la commande `ng build` ou `ng serve` :</span><span class="sxs-lookup"><span data-stu-id="6dcf3-168">To use the AOT compiler, add `--aot` to the `ng build` or `ng serve` command:</span></span>

```bash
ng build --aot
ng serve --aot
```

> [!NOTE]
> <span data-ttu-id="6dcf3-169">Pour en savoir plus sur le compilateur Ahead-of-Time (AOT) d’Angular, consultez le [guide officiel](https://angular.io/guide/aot-compiler).</span><span class="sxs-lookup"><span data-stu-id="6dcf3-169">To learn more about the Angular Ahead-of-Time (AOT) compiler, see the [official guide](https://angular.io/guide/aot-compiler).</span></span>
