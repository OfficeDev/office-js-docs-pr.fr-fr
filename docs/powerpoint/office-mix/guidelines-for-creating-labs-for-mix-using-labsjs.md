
# <a name="guidelines-for-creating-labs-for-mix-using-labsjs"></a>Lignes directrices relatives à la création d’ateliers pour Mix à l’aide de LabsJS



La bibliothèque LabsJS (labs.js) prend en charge l’écriture de Compléments Office (appelés « ateliers ») qui s’intègrent avec Office Mix. Office Mix, puis affiche les ateliers à l’aide de Microsoft PowerPoint. Même si nous nommons ces composants des « ateliers », nous créons bien des Compléments Office spéciaux qui sont des Compléments Office Mix.

Le contenu LabsJS vous permet d’implémenter l’API JavaScript labs.js en vous donnant des conseils et des exemples. Cette bibliothèque est créée au-dessus de l’ [Interface API JavaScript pour Office](http://dev.office.com/reference/add-ins/javascript-api-for-office) (Office.js) et fournit une couche d’abstraction optimisée pour les compléments incorporés dans Office Mix.


## <a name="general-guidelines"></a>Recommandations générales


Les sections suivantes correspondant aux recommandations générales relatives à l’écriture de compléments à l’aide de l’API LabJS.


### <a name="scripts"></a>Scripts

Étant donné que la bibliothèque labs.js est une couche d’abstraction sur office.js et que, par conséquent, elle comporte une dépendance sur office.js, les fichiers office.js et labs.js doivent être inclus dans vos projets de développement. 

Vous pouvez référencer la bibliothèque office.js à l’adresse suivante :  `<script src="https://sforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>`.

La bibliothèque labs.js est incluse avec le Kit de développement logiciel (SDK) LabsJS. Vous pouvez également référencer la bibliothèque labs.js sur un réseau de distribution de contenu (CDN) à l’adresse  <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>. La version de production de votre atelier doit faire référence à la version stockée sur le CDN.


 >**Remarque** :  en plus du fichier JavaScript (labs-1.0.4.js), nous vous fournissons un fichier de définition TypeScript de l’API labs (labs-1.0.4.d.ts). Le fichier de définition a été créé sur TypeScript, version 0.9.1.1.


### <a name="callbacks-and-error-handling"></a>Rappels et gestion des erreurs

Plusieurs méthodes fonctionnent de manière asynchrone dans l’API labs.js. Pour ces opérations, l’API adopte une interface de rappel standard :  **ILabCallback**. 


```js
function(err, result) {
}
```

La méthode de rappel prend deux paramètres :  _err_ et _result_. Le champ  _err_ conserve la valeur **null**, sauf en présence d’une erreur. Le champ  _result_ renvoie le résultat de l’opération.

L’opération de rappel ne se déclenche jamais immédiatement, même si le résultat est tout de suite disponible. Elle se déclenche sur une exécution distincte de la boucle d’événements JavaScript (via l’appel  **setTimeout**). En adoptant cette définition de rappel, vous pouvez facilement intégrer labs.js avec l’API de promesse de votre choix. Par exemple, vous pouvez remplacer les promesses jQuery pour ces rappels par une simple méthode de conversion, comme indiqué dans l’exemple suivant.




```js
function createCallback<T>(deferred: JQueryDeferred<T>): Labs.Core.ILabCallback<T> {
    return (err, data) => {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```


### <a name="lab-host-and-defaultlabhost"></a>Hôte de l’atelier et DefaultLabHost

L’hôte de l’atelier ( **ILabHost**) est le pilote sous-jacent qui prend en charge le développement des ateliers. Par défaut, il s’agit d’un hôte qui s’intègre avec office.js.

À des fins de test, et pour exécuter votre atelier dans le fichier labhost.html, vous devez passer à un hôte fonctionnant dans l’environnement de simulation. L’exemple de code suivant vous montre comment effectuer cette opération à l’aide d’un paramètre de requête. Vous pouvez également modifier  **DefaultHostBuilder** de façon à intégrer le complément de votre atelier avec une plateforme entièrement différente.




```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```


### <a name="initialization"></a>Initialisation

L’initialisation établit le chemin de communication entre l’atelier et son hôte. Initialisez votre atelier en appelant la méthode suivante :


```js
Labs.connect((err, connectionResponse) => {});
```

Une fois l’initialisation terminée, vous pouvez appeler d’autres méthodes de l’API labs.js. Le paramètre  _connectionResponse_ contient des informations sur l’hôte, l’utilisateur et la connexion. Pour plus d’informations sur les valeurs renvoyées, voir [Labs.Core.IConnectionResponse](http://dev.office.com/reference/add-ins/office-mix/labs.core.iconnectionresponse).


### <a name="time-format"></a>Format d’heure

Labs.js stocke les nombres sous forme de millisecondes écoulées depuis le 1er janvier 1970 UTC. Cela correspond au format de date de l’ [objet Date](http://msdn.microsoft.com/fr-fr/library/ie/cd9w2te4%28v=vs.94%29.aspx)JavaScript.


### <a name="timeline"></a>Chronologie

L’atelier peut également interagir avec la chronologie du lecteur de leçon. La chronologie permet à l’atelier de demander au lecteur de leçon de passer à la diapositive suivante. L’objet de chronologie est récupéré en appelant la méthode  **Labs.getTimeline**.


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="handling-events"></a>Gestion des événements


L’API d’événements LabsJS suit les événements propres à l’atelier et vous permet d’ajouter des gestionnaires d’événements, afin que vous puissiez agir sur les événements ou y répondre. Les méthodes d’événement, qui sont au nombre de trois, se trouvent sur l’objet  **EventTypes** :  **ModeChanged**,  **Activate** et **Deactivate**. 


### <a name="mode-change"></a>Changement de mode

L’événement  **ModeChanged** se déclenche lorsque l’atelier spécifié passe du mode Modification au mode Affichage. Le mode Modification est visible lorsque l’atelier est affiché dans le mode Modification de PowerPoint. Le mode Affichage est visible lorsque PowerPoint affiche le diaporama ou lorsque l’atelier est affiché dans le lecteur de leçon Office Mix. Le mode Affichage doit toujours afficher ce que l’utilisateur voit lorsqu’il commence l’atelier. Le mode Modification permet à l’utilisateur de configurer l’atelier.

Les données dans l’objet  **ModeChangedEventData**, qui est transmis au rappel, contient des informations sur le mode actuel. Le code suivant présente l’utilisation de l’événement  **ModeChanged**.




```js
Labs.on(Labs.Core.EventTypes.ModeChanged, (data) => {
    var modeChangedEvent = <Labs.Core.ModeChangedEventData> data;
    this.switchToMode(modeChangedEvent.mode);
});
```


### <a name="activate"></a>Activer

L’événement  **activate** se déclenche lorsque la diapositive PowerPoint sur laquelle se trouve l’atelier est activée dans le lecteur de leçon.


```js
Labs.on(Labs.Core.EventTypes.Activate, (data) => {
    //  is now on the active slide
});
```


### <a name="deactivate"></a>Désactiver

L’événement  **deactivate** se déclenche lorsque la diapositive PowerPoint sur laquelle se trouve l’atelier n’est plus active.


```js
Labs.on(Labs.Core.EventTypes.Deactivate, (data) => {                
    //  is no longer on the active slide
});
```


### <a name="timeline"></a>Chronologie

L’atelier peut également interagir avec la chronologie du lecteur de leçon. La chronologie permet à l’atelier de demander au lecteur de leçon de passer à la diapositive suivante. L’objet de chronologie est récupéré en appelant la méthode  **Labs.getTimeline**.


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="additional-resources"></a>Ressources supplémentaires



- [Compléments Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
