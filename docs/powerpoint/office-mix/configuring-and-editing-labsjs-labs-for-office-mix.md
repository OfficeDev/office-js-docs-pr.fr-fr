
# <a name="configuring-and-editing-labsjs-labs-for-office-mix"></a>Configuration et modification d’ateliers LabsJS pour Office Mix



Office Mix fournit des méthodes office.js pour obtenir et définir des configurations d’atelier. La configuration indique à Office Mix le type d’atelier que vous créez, ainsi que le type de données qui seront renvoyées par l’atelier. Ces informations sont utilisées pour collecter et visualiser des analyses.

## <a name="getting-the-lab-editor"></a>Obtention de l’éditeur d’atelier

L’éditeur d’atelier, l’objet [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md), vous permet de modifier votre atelier, ainsi que d’obtenir et de définir la configuration de ce dernier. Après avoir terminé de modifier votre atelier, vous devez appeler la méthode  **Done**. Toutefois, l’appel de la méthode  **Done** n’est pas requis, sauf lorsque vous essayez de récupérer ou d’exécuter un atelier que vous modifiez. Vous ne pouvez ouvrir qu’une seule instance de l’atelier à la fois.

Le code suivant vous montre comment obtenir l’éditeur d’atelier.




```js
Labs.editLab((err, labEditor) => {
    if (err) {
        handleError();
        return;
    }
    _labEditor = labEditor;
});
```

Utilisez les méthodes  **getConfiguration** et **setConfiguration** sur [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md) pour stocker la configuration d’un atelier spécifique. La configuration ([Labs.Core.IConfiguration](../../../reference/office-mix/labs.core.iconfiguration.md)) indique à Office Mix les données qui seront collectées et traitées par l’atelier. Une configuration contient des informations générales sur un atelier, notamment son nom, sa version et d’autres options de configuration. La définition des composants d’atelier est la partie la plus importante de la configuration.

Le code suivant illustre comment définir et obtenir une configuration. Pour définir une configuration, il vous suffit de créer l’objet de configuration et d’appeler ensuite la méthode  **setConfiguration**. Ensuite, pour récupérer la configuration, vous devez appeler la méthode  **getConfiguration** sur l’objet d’éditeur d’atelier.




```js

///////  Set the configuration /////

var activityComponent: Labs.Components.IActivityComponent = {
    type: Labs.Components.ActivityComponentType,
    name: uri,
    values: {},
    data: {
        uri: uri
    },
    secure: false
};
var configuration = {
    appVersion: { major: 1, minor: 1 },
    components: [activityComponent],
    name: configurationName,
    timeline: null,
    analytics: null
};
this._labEditor.setConfiguration(configuration, (err, unused) => { })

```




```js

///////  Get the configuration  //////

labEditor.getConfiguration((err, configuration) => {
});
```


## <a name="closing-the-editor"></a>Fermeture de l’éditeur

Pour fermer l’éditeur, appelez la méthode  **Done** sur l’éditeur lorsque vous avez terminé de modifier l’atelier. Vous ne pouvez pas exécuter et modifier un atelier en même temps. Cependant, après avoir appelé la méthode  **Done**, vous pouvez soit modifier, soit exécuter l’atelier.


## <a name="interacting-with-a-lab"></a>Interaction avec un atelier

Une fois la configuration de l’atelier définie, vous pouvez interagir avec l’atelier. Lorsque ce dernier est exécuté dans PowerPoint, les interactions sont simulées. Lorsqu’il est exécuté dans le lecteur de leçon Office Mix, les données sont stockées dans la base de données Office Mix et utilisées dans les analyses.


### <a name="getting-the-lab-instance"></a>Obtention de l’instance d’atelier

Vous interagissez avec l’atelier à l’aide de l’objet [Labs.LabInstance](../../../reference/office-mix/labs.labinstance.md), qui est une instance de l’atelier configuré pour l’utilisateur en cours. Pour exécuter l’atelier, appelez la fonction [Labs.takeLab](../../../reference/office-mix/labs.takelab.md).


```js
Labs.takeLab((err, labInstance) => {
    this._labInstance = labInstance;
    var activityComponentInstance = <Labs.Components.ActivityComponentInstance> this._labInstance.components[0];
    // populate the UI based on the instance    
});
```

L’objet d’instance contient un tableau des instances de composants ([Labs.ComponentInstanceBase](../../../reference/office-mix/labs.componentinstancebase.md), [Labs.ComponentInstance](../../../reference/office-mix/labs.componentinstance.md)) qui est mis en correspondance avec les composants indiqués dans la configuration. En fait, une instance est tout simplement une version transformée de la configuration qui est utilisée pour joindre des ID côté serveur à des objets d’instance, ainsi que pour dissimuler certains champs à l’utilisateur le cas échéant (par exemple, des conseils, des réponses, etc.).


### <a name="managing-state"></a>Gestion de l’état

L’état correspond au stockage temporaire associé à un utilisateur exécutant un atelier donné. Vous pouvez utiliser l’emplacement de stockage pour conserver des informations entre des appels successifs de l’atelier lab. Par exemple, un atelier de programmation pourrait stocker le travail en cours de l’utilisateur.

Pour  **set** l’état, utilisez le code suivant.




```js
labInstance.setState(this._labState(), (err, unused) => { 
    // If no error, state has successfully been stored by the host.
});
```

Pour  **get** l’état, utilisez le code suivant.




```js
labInstance.getState((err, state) => {
    // If no error, the state parameter contains the set state.
});
```


## <a name="component-instances-and-results"></a>Instances de composant et résultats

Vous trouverez ci-après une vue d’ensemble de la procédure d’implémentation des instances des quatre types de composants, ainsi que de courts exemples des méthodes de composant. 

Vous devez toutefois commencer par vous familiariser avec deux concepts essentiels lorsque vous utilisez des instances de composant. Le premier est le concept de  **tentatives** et de **valeurs**.

 **Tentatives**

Le terme « tentative » est employé lorsqu’un utilisateur essaye de terminer une instance de composant. Par exemple, dans le cas d’une question à choix multiple, une tentative démarre lorsque l’utilisateur commence à étudier le problème et se termine lorsqu’une note finale est attribuée. Les analyses Office Mix regroupent ensuite les résultats de l’utilisateur pour le problème.


 >**Remarque** : Les tentatives peuvent être utilisées pour tous les types de composants, sauf le type **DynamicComponent**.

Vous pouvez récupérer les résultats de l’ensemble des tentatives associées à une instance de composant donnée à l’aide de la méthode  **getAttempts**. Une fois les résultats récupérés, l’utilisateur peut soit réessayer l’une des tentatives existantes à l’aide de la méthode  **resume**, soit créer une tentative à l’aide de la méthode  **createAttempt**. Le processus est illustré dans l’exemple suivant.




```js
var attemptsDeferred = $.Deferred();
activityComponentInstance.getAttempts(createCallback(attemptsDeferred));
var attemptP = attemptsDeferred.promise().then((attempts) => {
    var currentAttemptDeferred = $.Deferred();
    if (attempts.length > 0) {
        currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
    } else {
        activityComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
    }
    return currentAttemptDeferred.then((currentAttempt: Labs.Components.ActivityComponentAttempt) => {
        var resumeDeferred = $.Deferred();
        currentAttempt.resume(createCallback(resumeDeferred));
        return resumeDeferred.promise().then(() => {
            return currentAttempt;
        });
    });
});
```

 **Valeurs**

Les instances de composant contiennent un dictionnaire de clés mis en correspondance avec un tableau de valeurs. Vous pouvez utiliser le tableau pour stocker des conseils, des commentaires ou tout autre ensemble de valeurs à associer au composant. L’instance de composant vous donne accès à ces valeurs à l’aide de la méthode  **getValues**.

Par exemple, si l’utilisateur lance une requête pour obtenir une valeur de conseil, cette information est indiquée dans les analyses. Les valeurs sont suivies à chaque tentative.

L’exemple de code suivant montre comment lancer une requête de conseil.




```js
// Take a hint.
var hints = attempt.getValues("hints");
hints[0].getValue((err, hint) => {
    // If no error, hint param will contain the hint data.
});
```


### <a name="activitycomponentinstance"></a>ActivityComponentInstance


Utilisez l’objet  **ActivityComponentInstace** pour suivre l’interaction entre un utilisateur et un composant d’activité. Cette classe fournit une méthode  **complete** permettant d’indiquer que l’utilisateur a fini d’interagir avec l’activité. La méthode peut indiquer que l’utilisateur a terminé une tâche qui lui était attribuée, une lecture ou toute autre tâche associée à l’activité. Le code suivant montre comment utiliser la méthode  **complete**.


```js
attempt.complete((err, unused) => { 
    // Called after the host has stored the completion.
});
```


### <a name="choicecomponentinstance"></a>ChoiceComponentInstance


Utilisez l’objet  **ChoiceComponentInstance** pour suivre l’interaction entre un utilisateur et un composant de choix. Les composants de choix correspondent à des problèmes pour lesquels l’utilisateur doit faire un choix parmi une liste donnée. Il peut y avoir une réponse correcte ou non. La classe fournit deux méthodes principales : **getSubmissions** et **submit**. La méthode  **getSubmissions** vous permet de récupérer les envois stockés précédemment et la méthode  **submit** permet le stockage d’un nouvel envoi. Le code suivant illustre l’utilisation des méthodes.


```js
///  using getSubmission method  ///
var submissions = this._attempt.getSubmissions();
```


```js
///  using submit method  ///
this._attempt.submit(
    new Labs.Components.ChoiceComponentAnswer(submission), 
    new Labs.Components.ChoiceComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### <a name="inputcomponentinstance"></a>InputComponentInstance


Utilisez l’objet  **InputComponentInstance** pour suivre l’interaction entre un objet et un composant de saisie. La classe fournit deux méthodes principales : **getSubmission** et **submit**. La méthode  **getSubmissions** vous permet de récupérer des envois stockés précédemment et la méthode  **submit** vous permet de stocker un nouvel envoi. L’extrait de code suivant illustre l’utilisation de la méthode  **getSubmissions**.


```js
var submissions = this._attempt.getSubmissions();
```

Lorsque vous utilisez la méthode  **submit**, l’objet  **InputComponentAnswer** représente la réponse envoyée et l’objet  **InputComponentResult** contient le résultat. La valeur renvoyée est un objet  **InputComponentSubmission** qui contient la réponse, le résultat et un horodateur, qui indique la date/l’heure d’envoi du résultat.




```js
this._attempt.submit(
    new Labs.Components.InputComponentAnswer(submission), 
    new Labs.Components.InputComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### <a name="dynamiccomponentinstance"></a>DynamicComponentInstance


Utilisez l’objet  **DynamicComponentInstance** pour suivre l’interaction entre un utilisateur et un composant dynamique. Les méthodes principales dans cette classe sont les suivantes : **getComponents**,  **createComponent** et **close**.

La méthode  **getComponents** vous permet de récupérer la liste des instances de composant créées précédemment, comme indiqué dans l’exemple suivant.




```js
dynamicComponentInstance.getComponents((err, components) => {
    // Upon success, components contains a list of previously created component instances.
});
```

La méthode  **createComponent** crée un composant et renvoie cette instance de composant, comme indiqué dans l’exemple suivant.




```js
var inputComponentHints = [];
for (var i = 0; i < data.hints.length; i++) {
    inputComponentHints.push({
        isHint: true,
        value: data.hints[i]        
    });
}
var inputComponent = {
    maxScore: 1,
    timeLimit: 0,
    hasAnswer: true,
    answer: data.answerData.solution,
    type: Labs.Components.InputComponentType,
    name: data.name,
    values: { hints: inputComponentHints },
    secure: false
};
var currentAttemptDeferred = $.Deferred();
var dynamicComponent = labInstance.components[0];
dynamicComponent.createComponent(inputComponent, function(err, inputComponentInstance) {
    // Create will return the instance for the specified component.
})
```

Utilisez la méthode  **close** pour indiquer que vous avez terminé d’utiliser le composant dynamique pour créer des composants. Vous pouvez également utiliser une méthode booléenne  **isClosed** pour vérifier si l’instance de composant dynamique a été fermée. Le code suivant illustre l’utilisation de la méthode  **close**.




```js
dynamicComponentInstance.close((err, unused) => {
    // Called after the server has processed the close attempt.
});
```


## <a name="additional-resources"></a>Ressources supplémentaires



- [Compléments Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Procédure pas à pas : Création de votre premier laboratoire pour Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
