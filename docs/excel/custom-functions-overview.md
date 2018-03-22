---
title: Créer des fonctions personnalisées dans Excel (Aperçu)
description: ''
ms.date: 01/23/2018
---

# <a name="create-custom-functions-in-excel-preview"></a>Créer des fonctions personnalisées dans Excel (Aperçu)

Les fonctions personnalisées (similaires aux fonctions définies par l’utilisateur) permettent aux développeurs d’ajouter n’importe quelle fonction JavaScript à Excel en utilisant un complément. Les utilisateurs peuvent accéder aux fonctions personnalisées comme à toute autre fonction native dans Excel (par exemple, =SUM()). Cet article explique comment créer des fonctions personnalisées dans Excel.

Voici à quoi ressemblent les fonctions personnalisées dans Excel :

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

Voici le code d’un exemple de fonction personnalisée qui ajoute 42 à une paire de nombres.

```js
function add42 (a, b) {
    return a + b + 42;
}
```

Les fonctions personnalisées sont désormais disponibles en version d’évaluation. Pour les tester, procédez comme suit :

1.  Rejoignez le programme [Office Insider](https://products.office.com/fr-fr/office-insider) pour installer la version d’Excel 2016 requise pour les fonctions personnalisées sur votre ordinateur (version 16.8711 ou ultérieure). Vous devez choisir le canal « Insider » pour obtenir l’aperçu des fonctions personnalisées à utiliser.
2.  Clonez le référentiel [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) et suivez les instructions dans *README.md* afin de lancer le complément dans Excel.
3.  Saisissez `=CONTOSO.ADD42(1,2)` dans une cellule, puis appuyez sur **Entrée** pour exécuter la fonction personnalisée.
4.  Si vous avez des questions, posez-les sur Office Insider à l’aide de la balise [office-js](https://stackoverflow.com/questions/tagged/office-js).

Reportez-vous à la section Problèmes connus à la fin de cet article. Elle inclut les limites actuelles des fonctions personnalisées et sera mise à jour au fil du temps.

## <a name="learn-the-basics"></a>Notions fondamentales


Dans le référentiel d’exemple cloné, vous trouverez les fichiers suivants :

-   *customfunctions.js*, qui contient les éléments suivants :

    -   Code de fonction personnalisée à ajouter à Excel.
    -   Code d’enregistrement pour connecter votre fonction personnalisée à Excel. Avec l’enregistrement, vos fonctions personnalisées apparaissent dans la liste des fonctions disponibles affichée lorsque les utilisateurs saisissent du texte dans les cellules.
-   *customfunctions.html*, qui indique une référence de &lt;script&gt; à *customfunctions.js*. Ce fichier n’affiche pas d’interface utilisateur dans Excel.
-   *manifest.xml*, qui indique à Excel l’emplacement de vos fichiers HTML et JS nécessaires à l’exécution des fonctions personnalisées.

### <a name="javascript-file-customfunctionsjs"></a>Fichier JavaScript (*customfunctions.js*)

Le code suivant dans customfunctions.js déclare la fonction personnalisée `add42`, puis enregistre la fonction dans Excel.

```js
function add42 (a, b) {
    return a + b + 42;
}

Excel.Script.customFunctions["CONTOSO"]["ADD42"] = {
    call: add42,
    description: "Adds 42 to the sum of two numbers",
    helpUrl: "https://www.contoso.com/help.html",
    result: {
        resultType: Excel.CustomFunctionValueType.number,
        resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    parameters: [{
        name: "num 1",
        description: "The first number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    {
        name: "num 2",
        description: "The second number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    }],
    options:{ batch: false, stream: false }
};

Excel.run(function(ctx) {
    ctx.workbook.customFunctions.addAll();
});
```

**L’enregistrement** de la fonction personnalisée utilise le bloc de code `Excel.Script.customFunctions["CONTOSO"]["ADD42"]`. Vous avez besoin des paramètres suivants pour enregistrer la fonction dans Excel :

-   Nom de la fonction et préfixe : la première valeur dans `Excel.Script.customFunctions` est le préfixe (dans ce cas, CONTOSO est le préfixe). La seconde valeur dans `Excel.Script.customFunctions` est le nom de la fonction (dans ce cas, ADD42 est le nom de la fonction). Dans Excel, le préfixe et le nom de la fonction sont séparés par un point : pour utiliser votre fonction personnalisée, associez le préfixe de la fonction (CONTOSO) au nom de la fonction (ADD42), et entrez `=CONTOSO.ADD42` dans une cellule. Par convention, les préfixes et les noms de fonction contiennent des lettres majuscules. Le préfixe est destiné à être utilisé comme identificateur de votre complément.
-   `call` : définit la fonction JavaScript à appeler (par exemple, `add42`). Le nom de la fonction JavaScript ne doit pas forcément correspondre au nom que vous avez enregistré dans Excel.
-   `description` : la description apparaît dans le menu de saisie semi-automatique dans Excel.
-   `helpUrl` : lorsque l’utilisateur demande de l’aide concernant une fonction, Excel ouvre un volet Office et affiche la page web accessible via cette URL.
-   `result` : Définit le type d’informations renvoyées par la fonction à Excel.

    -   `resultType` : votre fonction peut renvoyer une valeur `"string"` ou `"number"` (également utilisées pour les dates et les devises). Pour plus d’informations, reportez-vous à la section [Énumérations des fonctions personnalisées](https://dev.office.com/reference/add-ins/excel/customfunctionsenumerations).
    -   `resultDimensionality` : votre fonction peut renvoyer une valeur (`"scalar"`) simple ou une `"matrix"` de valeurs. Dans le cas d’une matrice de valeurs, la fonction renvoie un tableau, où chaque élément de tableau est un autre tableau qui représente une ligne de valeurs. Pour plus d’informations, reportez-vous à la section [Énumérations des fonctions personnalisées](https://dev.office.com/reference/add-ins/excel/customfunctionsenumerations). L’exemple suivant renvoie une matrice de valeurs à 3 lignes et 2 colonnes à partir d’une fonction personnalisée.

        ```js
        return [["first","row"],["second","row"],["third","row"]];
        ```

-   Votre fonction personnalisée accepte les arguments comme entrées. Les arguments transmis à votre fonction personnalisée sont spécifiés dans la propriété des *paramètres*. L’ordre des paramètres dans la définition doit correspondre à l’ordre dans la fonction JavaScript. Pour chaque paramètre, définissez ces propriétés :

    -   `name` : chaîne affichée dans Excel pour représenter le paramètre.
    -   `description` : chaîne affichée pour obtenir plus d’informations sur le paramètre.
    -   `valueType` : valeur `"number"` ou `"string"`, de même que pour la propriété resultType décrite précédemment.
    -   `valueDimensionality` : valeur `"scalar"` ou `"matrix"` de valeurs, de même que pour la propriété resultDimensionality décrite précédemment. Les paramètres de type matrice permettent à l’utilisateur de sélectionner des plages plus volumineuses qu’une seule cellule.

-   `options` : permet d’activer des types spéciaux de fonctions personnalisées qui sont décrits de manière plus détaillée par la suite dans cet article.

Pour terminer l’enregistrement de toutes les fonctions définies en utilisant `Excel.Script.customFunctions`, appelez `CustomFunctions.addAll()`.

Après l’enregistrement, les fonctions personnalisées sont disponibles pour l’utilisateur dans tous les classeurs (pas seulement dans celui où le complément a été exécuté initialement). Les fonctions sont affichées dans le menu de saisie semi-automatique lorsque l’utilisateur commence à saisir leur nom. Au cours du développement et du test, vous pouvez vider manuellement le cache de votre ordinateur de métadonnées d’inscription en supprimant le dossier `<user>\AppData\Local\Microsoft\Office\16.0\Wef\CustomFunctions`.


### <a name="manifest-file-manifestxml"></a>Fichier manifeste (*manifest.xml*)

L’exemple suivant dans manifest.xml permet à Excel de rechercher le code de vos fonctions.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">

    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="scriptURL" />
                        <!— Required. The Developer Preview does not use the Script element.-->
                    </Script>
                    <Page>
                        <SourceLocation resid="pageURL"/>
                    </Page>
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>

    <Resources>
        <bt:Urls>
            <bt:Url id="scriptURL" DefaultValue="https://www.contoso.com/addin/customfunctions.js" />
            <bt:Url id="pageURL" DefaultValue="https://www.contoso.com/addin/customfunctions.html" />
        </bt:Urls>
    </Resources>

</VersionOverrides>

```

Le code ci-dessus spécifie ce qui suit :

-   Élément `<Script>` : obligatoire mais non utilisé dans la version d’évaluation pour développeurs.
-   Élément `<Page>` : lien vers la page HTML de votre complément. La page HTML inclut une référence de &lt;script&gt; au fichier JavaScript (*customfunctions.js*) qui contient la fonction personnalisée et le code d’enregistrement. La page HTML est une page masquée qui n’est jamais affichée dans l’interface utilisateur.

## <a name="asynchronous-functions"></a>Fonctions asynchrones

Si votre fonction personnalisée récupère des données à partir du web, vous devez effectuer un appel asynchrone pour les extraire. Lors de l’appel à des services web externes, votre fonction personnalisée doit :

1.   Renvoyer une promesse JavaScript à Excel
2.   Vérifier la demande http pour appeler le service externe
3.   Résoudre la promesse à l’aide du rappel `setResult` `setResult` envoie la valeur à Excel.

Le code suivant indique un exemple de fonction personnalisée qui récupère la température d’un thermomètre.

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult, setError){
        sendWebRequestExample(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a>Fonctions de flux

Les fonctions personnalisées de flux vous permettent d’afficher des données dans des cellules à plusieurs reprises au fil du temps, sans devoir attendre qu’Excel ou que des utilisateurs demandent à effectuer le calcul à nouveau. Par exemple, la fonction personnalisée `incrementValue` dans le code suivant ajoute un nombre au résultat à chaque seconde qui passe, et Excel affiche automatiquement chaque nouvelle valeur à l’aide du rappel `setResult`. Pour afficher le code d’enregistrement utilisé avec `incrementValue`, consultez le fichier *customfunctions.js*.

```js
function incrementValue(increment, caller){ 
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

Pour les fonctions de flux, le paramètre final, `caller`, n’est jamais spécifié dans votre code d’enregistrement et ne s’affiche pas dans le menu de saisie semi-automatique pour les utilisateurs d’Excel lorsqu’ils entrent la fonction. Il s’agit d’un objet contenant une fonction de rappel `setResult` utilisée pour transmettre des données de la fonction à Excel afin de mette à jour la valeur d’une cellule. Afin qu’Excel transmette la fonction `setResult` dans l’objet `caller`, vous devez déclarer la prise en charge de la diffusion en continu lors de l’enregistrement de la fonction en définissant le paramètre `stream` sur `true`.

## <a name="cancellation"></a>Annulation

Vous pouvez annuler les fonctions de flux et les fonctions asynchrones. L’annulation de vos appels de fonction permet de considérablement réduire leur consommation de bande passante, la mémoire de travail et la charge de l’UC. Excel annule les appels de fonction dans les situations suivantes :
- L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.
- Un des arguments (entrées) de la fonction est modifié. Dans ce cas, un nouvel appel de fonction est déclenché en plus de l’annulation.
- L’utilisateur déclenche le nouveau processus de calcul manuellement. Comme pour le cas précédent, un nouvel appel de fonction est déclenché en plus de l’annulation.

Le code suivant affiche l’exemple précédent avec l’annulation mise en œuvre. Dans le code, l’objet `caller` contient une propriété `onCanceled` qui doit être définie pour chaque fonction personnalisée. Pour qu’Excel appelle la fonction `onCanceled`, vous devez déclarer la prise en charge de l’annulation lors de l’enregistrement de votre fonction en définissant le paramètre `cancelable` sur `true`.

```js
function incrementValue(increment, caller){ 
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);

    caller.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-state"></a>État enregistré

Les fonctions personnalisées peuvent enregistrer des données dans des variables JavaScript globales. Lors d’appels ultérieurs, votre fonction personnalisée peut utiliser les valeurs enregistrées dans ces variables. L’état enregistré est utile lorsque les utilisateurs entrent plusieurs instances de la même fonction personnalisée, car ils doivent partager des données entre eux. Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter d’effectuer des appels supplémentaires à la même ressource web.

Le code suivant illustre une implémentation de la fonction de flux précédente relative à la température qui enregistre l’état à l’aide de la variable `savedTemperatures`. Le code montre les concepts suivants :

-   **Enregistrement des données.** `refreshTemperature` est une fonction de flux qui lit la température d’un thermomètre spécifique à chaque seconde qui passe. Les nouvelle températures sont enregistrées dans la variable savedTemperatures.

-   **Utilisation des données enregistrées.** `streamTemperature` met à jour les valeurs de température affichées dans l’interface utilisateur Excel à chaque seconde. Les températures sont lues à partir de `savedTemperature`, puis envoyées à l’interface utilisateur Excel en utilisant `setResult`. Les utilisateurs peuvent appeler `streamTemperature` à partir de plusieurs cellules dans l’interface utilisateur Excel. Chaque appel à `streamTemperature` entraîne la lecture des données à partir de `savedTemperatures`.

> Dans ce cas, nous enregistrons `streamTemperature` en tant que fonction personnalisée dans Excel.

```js
var savedTemperatures{};

function streamTemperature(thermometerID, caller){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID);
     }

     function getNextTemperature(){
         caller.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
     }
     getNextTemperature();
}

function refreshTemperature(thermometerID){
     sendWebRequestExample(thermometerID, function(data){
         savedTemperatures[thermometerID] = data.temperature;
     });
     setTimeout(function(){
         refreshTemperature(thermometerID);
     }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="working-with-ranges-of-data"></a>Utilisation des plages de données

Votre fonction personnalisée accepte les plages de données en tant que paramètres. Sinon, vous pouvez renvoyer une plage de données à partir d’une fonction personnalisée.

Par exemple, supposons que votre fonction renvoie la deuxième température la plus élevée à partir d’une plage de valeurs de température stockées dans Excel. La fonction suivante prend le paramètre `temperatures`, c’est-à-dire un type de paramètre `Excel.CustomFunctionDimensionality.matrix`.

```js
function secondHighestTemp(temperatures){ 
     var highest = -273, secondHighest = -273;
     for(var i = 0; i < temperatures.length; i++){
         for(var j = 0; j < temperatures[i].length; j++){
             if(temperatures[i][j] <= highest){
                 secondHighest = highest;
                 highest = temperatures[i][j];
             }
             else if(temperatures[i][j] <= secondHighest){
                 secondHighest = temperatures[i][j];
             }
         }
     }
     return secondHighest;
 }
```

Si vous créez une fonction qui renvoie une plage de données, il est nécessaire d’indiquer une formule de tableau dans Excel pour identifier la plage complète de valeurs. Pour plus d’informations, reportez-vous à [Instructions et exemples de formules matricielles](https://support.office.com/fr-fr/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7).

## <a name="known-issues"></a>Problèmes connus

Les fonctionnalités suivantes ne sont pas encore prises en charge dans la version d’évaluation pour développeurs.

-   Traitement par lots : vous permet d’agréger plusieurs appels à la même fonction pour améliorer les performances.

-   Les descriptions de paramètre et les URL d’aide ne sont pas encore utilisées par Excel.

-   Le déploiement de compléments utilisant des fonctions personnalisées sur AppSource ou via Office 365 a centralisé le déploiement.

-   Les fonctions personnalisées ne sont pas disponibles dans Excel pour Mac, Excel pour iOS et Excel Online.

-   Actuellement, les compléments s’appuient sur un processus de navigateur masqué pour exécuter les fonctions personnalisées. À l’avenir, JavaScript s’exécutera directement sur certaines plateformes pour garantir que les fonctions personnalisées sont plus rapides et utilisent moins de mémoire. Par ailleurs, la page HTML référencée par l’élément &lt;Page&gt; dans le fichier manifeste ne sera pas nécessaire pour la plupart des plateformes, car Excel exécutera directement le code JavaScript. Pour vous préparer à ce changement, vérifiez que vos fonctions personnalisées n’utilisent pas le DOM de page web.

## <a name="changelog"></a>Journal des modifications

- **7 novembre 2017 :** mise à disposition des exemples et de la version d’évaluation des fonctions personnalisées
- **20 novembre 2017 :** correction du bogue de compatibilité pour les utilisateurs de la version 8801 et ultérieure
- **28 novembre 2017 :** prise en charge de l’annulation sur des fonctions asynchrones (nécessite la modification des fonctions de flux)
