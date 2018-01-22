
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul

L’objet [Document](http://dev.office.com/reference/add-ins/shared/document) expose des méthodes qui vous permettent de lire et d’écrire dans la sélection active de l’utilisateur dans un document ou une feuille de calcul. Pour cela, l’objet **Document** fournit les méthodes **getSelectedDataAsync** et **setSelectedDataAsync**. Cette rubrique explique comment lire, écrire et créer des gestionnaires d’événements pour détecter les changements intervenant dans la sélection de l’utilisateur.

La méthode **getSelectedDataAsync** ne fonctionne que sur la sélection active de l’utilisateur. Si vous devez conserver la sélection dans le document, afin que la même sélection soit disponible en lecture et en écriture dans les sessions exécutant votre complément, vous devez ajouter une liaison à l’aide de la méthode [Bindings.addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155.aspx) (ou créer une liaison à l’aide de l’une des autres méthodes « addFrom » de l’objet [Bindings](http://msdn.microsoft.com/en-us/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1.aspx)). Pour plus d’informations sur la création d’une liaison vers une zone d’un document et sur la lecture et l’écriture dans une liaison, voir [Liaisons de régions dans un document ou une feuille de calcul](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Lecture de données sélectionnées


L’exemple suivant montre comment obtenir les données d’une sélection dans un document en utilisant la méthode [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md).


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Dans cet exemple, le premier paramètre _coercionType_ est spécifié comme **Office.CoercionType.Text** (vous pouvez également spécifier ce paramètre en utilisant la chaîne littérale `"text"`). Cela signifie que la propriété [value](../../reference/shared/asyncresult.status.md) de l’objet [AsyncResult](http://dev.office.com/reference/add-ins/shared/asyncresult) qui est disponible à partir du paramètre _asyncResult_ dans la fonction de rappel renverra une **string** qui contient le texte sélectionné dans le document. La spécification de différents types de forçage de type produit des valeurs différentes. [Office.CoercionType](http://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) est une énumération des valeurs de types de forçage de type disponibles. **Office.CoercionType.Text** prend la valeur de la chaîne « text ».


 >**Conseil :**   **Quand utiliser le type de forçage (coercionType) de matrice (matrix) et de tableau (table) pour l’accès aux données ?** Si vous avez besoin que vos données tabulaires sélectionnées s’élargissent dynamiquement lorsque des lignes et des colonnes sont ajoutées, et que vous devez utiliser des en-têtes de tableau, vous devez utiliser le type de données de tableau (table) (en définissant le paramètre _coercionType_ de la méthode **getSelectedDataAsync** sur `"table"` ou **Office.CoercionType.Table**). L’ajout de lignes et de colonnes dans la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes est pris en charge uniquement pour les données de tableau. Si vous ne prévoyez pas d’ajouter des lignes et des colonnes, et que vos données ne nécessitent pas la fonctionnalité d’en-tête, vous devez utiliser le type de données de matrice (en définissant le paramètre _coercionType_ de la méthode **getSelecteDataAsync** sur `"matrix"` ou **Office.CoercionType.Matrix**), ce qui fournit un modèle d’interaction avec les données plus simple.

La fonction anonyme qui est transmise dans la fonction comme deuxième paramètre _callback_ est exécutée lorsque l’opération **getSelectedDataAsync** est terminée. La fonction est appelée avec un seul paramètre, _asyncResult_, qui contient le résultat et l’état de l’appel. Si l’appel échoue, la propriété [error](../../reference/shared/asyncresult.context.md) de l’objet **AsyncResult** donne accès à l’objet [Error](http://dev.office.com/reference/add-ins/shared/error). Vous pouvez vérifier la valeur des propriétés [Error.name](../../reference/shared/error.name.md) et [Error.message](../../reference/shared/error.message.md) pour déterminer les raisons de l’échec de l’opération. Sinon, le texte sélectionné dans le document s’affiche.

La propriété [AsyncResult.status](../../reference/shared/asyncresult.error.md) est utilisée dans l’instruction **if** pour tester la réussite de l’appel. [Office.AsyncResultStatus](http://dev.office.com/reference/add-ins/shared/asyncresultstatus-enumeration) est une énumération des valeurs de propriété **AsyncResult.status** disponibles. **Office.AsyncResultStatus.Failed** prend la valeur de la chaîne « failed » (et, de nouveau, peut également être spécifié comme chaîne littérale).


## <a name="write-data-to-the-selection"></a>Écriture de données dans la sélection


L’exemple suivant montre comment définir la sélection pour afficher « Hello World! ».


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Le passage de différents types d’objets pour le paramètre  _data_ produit différents résultats. Le résultat varie en fonction de la sélection actuelle dans le document, de l’application qui héberge votre complément, et de l’éventuel passage forcé des données dans la sélection actuelle.

La fonction anonyme transmise dans la méthode [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) comme paramètre _callback_ est exécutée quand l’appel anonyme est terminé. Lorsque vous écrivez des données dans la sélection à l’aide de la méthode **setSelectedDataAsync**, le paramètre _asyncResult_ du rappel donne uniquement accès à l’état de l’appel et à l’objet [Error](http://dev.office.com/reference/add-ins/shared/error) si l’appel échoue.

 **Remarque :** depuis la publication d’Excel 2013 SP1 et de la version correspondante d’Excel Online, vous pouvez désormais [définir la mise en forme lors de l’écriture d’un tableau sur la sélection active](../../docs/excel/format-tables-in-add-ins-for-excel.md).


## <a name="detect-changes-in-the-selection"></a>Détection de modifications dans la sélection


L’exemple suivant montre comment détecter des modifications dans la sélection à l’aide de la méthode [Document.addHandlerAsync](../../reference/shared/document.addhandlerasync.md) permettant d’ajouter un gestionnaire d’événements pour l’événement [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) sur le document.


```
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Le premier paramètre  _eventType_ spécifie le nom de l’événement auquel souscrire. Transmettre la chaîne `"documentSelectionChanged"` pour ce paramètre revient à transmettre le type d’événement **Office.EventType.DocumentSelectionChanged** de l’énumération [Office.EventType](http://dev.office.com/reference/add-ins/shared/eventtype-enumeration).

La fonction `myHander()` transmise dans la fonction comme deuxième paramètre _handler_ est un gestionnaire d’événements qui est exécuté lorsque la sélection change dans le document. La fonction est appelée avec un seul paramètre, _eventArgs_, qui contient une référence à un objet [DocumentSelectionChangedEventArgs](../../reference/shared/document.selectionchangedeventargs.md) quand l’opération asynchrone se termine. Vous pouvez utiliser la propriété [DocumentSelectionChangedEventArgs.document](../../reference/shared/document.selectionchangedeventargs.document.md) pour accéder au document qui a déclenché l’événement.


 >**Remarque**  Vous pouvez ajouter plusieurs gestionnaires d’événements pour un événement donné en rappelant la méthode  **addHandlerAsync** et en passant une fonction de gestionnaire d’événements supplémentaire pour le paramètre _handler_. Cela fonctionnera correctement à condition que le nom de chaque fonction de gestionnaire d’événements soit unique.


## <a name="stop-detecting-changes-in-the-selection"></a>Arrêt de la détection de modifications dans la sélection


L’exemple suivant montre comment arrêter l’écoute de l’événement [Document.SelectionChanged](../../reference/shared/document.selectionchanged.event.md) en appelant la méthode [document.removeHandlerAsync](../../reference/shared/document.removehandlerasync.md).


```
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Le nom de la fonction  `myHandler` passé en tant que deuxième paramètre _handler_ désigne le gestionnaire d’événements qui sera supprimé de l’événement **SelectionChanged**.


 >**Important :**  Si le paramètre facultatif  _handler_ est omis lors de l’appel à la méthode **removeHandlerAsync**, tous les gestionnaires d’événements du paramètre  _eventType_ spécifié sont supprimés.

