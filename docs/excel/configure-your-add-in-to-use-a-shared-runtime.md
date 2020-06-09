---
ms.date: 05/17/2020
title: Configurer votre complément Excel pour qu’il partage le runtime du navigateur
ms.prod: excel
description: Configurez votre complément Excel pour partager le runtime du navigateur et exécuter le ruban, le volet des tâches et le code de fonction personnalisée dans le même runtime.
localization_priority: Priority
ms.openlocfilehash: 8c16642f5a945e6156fcfd93c8b4cc088b616102
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609344"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime"></a>Configurer votre complément Excel pour utiliser un Runtime JavaScript partagé

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Lors de l’exécution d’Excel sur Windows ou Mac, votre complément exécute le code des boutons du ruban, des fonctions personnalisées et du volet des tâches dans des environnements runtime JavaScript distincts. Cela crée des limitations telles que le fait de ne pas pouvoir partager facilement des données globales, et n’ayant pas accès à toutes les fonctionnalités CORS à partir d’une fonction personnalisée.

Vous pouvez toutefois configurer votre complément Excel pour partager un code dans un runtime JavaScript partagé. Vous pouvez ainsi améliorer la coordination dans votre complément et accéder aux DOM et CORS à partir de toutes les parties de votre complément. Il vous permet également d’exécuter un code lorsque le document s’ouvre ou pendant la fermeture du volet des tâches. Si vous voulez configurer votre complément pour utiliser un runtime partagé, suivez les instructions contenues dans cet article.

## <a name="create-the-add-in-project"></a>Création du projet de complément

Si vous démarrez un nouveau projet, suivez ces étapes pour utiliser le générateur Yeoman et créer le projet de complément Excel. Exécutez la commande suivante, puis répondez aux invites avec les réponses suivantes :

```command line
yo office
```

- Choose a project type (Choisissez un type de projet) : **projet de complément Fonctions personnalisées Excel**
- Choose a script type (Choisissez un type de script) :  **JavaScript**
- What do you want to name your add-in? (Comment souhaitez-vous nommer votre complément ?)  **My Office Add-in**

![Capture d’écran de réponse aux invites à partir d’Office pour créer le projet de complément.](../images/yo-office-excel-project.png)

Après avoir exécuté l’Assistant, le générateur crée le projet et installe les composants Node de prise en charge.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Procédez comme suit pour configurer un projet nouveau ou existant de manière à utiliser un runtime partagé.

1. Démarrez Visual Studio Code et ouvrez le projet **My Office Add-in**.
2. Ouvrez le fichier **manifest.xml**.
3. Recherchez la section `<VersionOverrides>`, puis ajoutez l'exemple d'entrée suivante à la section `<Runtimes>`. La durée de vie doit être **longue** afin que les fonctions personnalisées puissent continuer de fonctionner même quand le volet Office est fermé. L'ID de ressources est `ContosoAddin.Url`, faisant par la suite référence à une chaîne dans la section des ressources. Vous pouvez utiliser n’importe quelle valeur d'ID de ressources souhaitée. Elle doit cependant correspondre à l'ID de ressources des autres éléments contenus dans les parties de votre complément.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. Dans l’élément `<Page>`, remplacez l’emplacement de la source **Functions.Page.Url** par **ContosoAddin.Url**. Cet ID de ressources correspond à l'ID de ressources de `<Runtime>`. Veuillez noter que si vous ne disposez pas de fonctions personnalisées, vous n’aurez pas d'entrée de **Page**. Vous pouvez par conséquent ignorer cette étape.

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. Dans la section `<DesktopFormFactor>`, changez la valeur **FunctionFile** de **Command.Url** pour utiliser **ContosoAddin.Url**. Veuillez noter que si vous n'avez pas de commandes d'action, vous ne disposerez pas d'entrée **FunctionFile**. Vous pouvez par conséquent ignorer cette étape.

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. Dans la section `<Action>`, remplacez l’emplacement de la source **Taskpane.Url** par **ContosoAddin.Url**. Veuillez noter que si vous n'avez pas de volet des tâches, vous ne disposerez pas de l'action **ShowTaskPane**. Vous pouvez par conséquent ignorer cette étape.

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. Ajoutez un nouvel **ID d’URL** pour **ContosoAddin.Url** pointant vers **taskpane.html**.

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. Enregistrez vos changements et reconstruisez le projet.

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a>Durée de vie de l’exécution

Lorsque vous ajoutez l’élément `Runtime`, vous spécifiez également une durée de vie ayant une valeur de `long` ou de `short`. Configurez cette valeur sur `long` pour tirer parti de fonctionnalités telles que le démarrage de votre complément lorsque le document s’ouvre, continuer à exécuter un code après la fermeture du volet des tâches, ou utiliser CORS et DOM à partir de fonctions personnalisées.

>! Note La valeur de la durée de vie par défaut est `short` , mais nous vous recommandons de l’utiliser `long` dans des compléments Excel. Si vous définissez votre Runtime sur `short` dans cet exemple, votre complément Excel démarre lorsque l’utilisateur appuie sur l’un de vos boutons du ruban, mais il peut se fermer une fois l’exécution de votre gestionnaire de ruban terminée. De la même façon, le complément démarre lorsque le volet des tâches est ouvert, mais il se peut se fermer à la fermeture du volet des tâches.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="multiple-task-panes"></a>Multiples volets des tâches

Ne concevez pas votre complément de sorte qu’il utilise plusieurs volets de tâches si vous envisagez d’utiliser un runtime partagé. Une exécution partagée prend uniquement en charge l’utilisation d’un volet de tâches. Notez que tout volet des tâches sans `<TaskpaneID>` est considéré comme un volet des tâches différent.

## <a name="next-steps"></a>Étapes suivantes

- Lisez l’article [Appeler des API Microsoft Excel à partir d’une fonction personnalisée](call-excel-apis-from-custom-function.md) pour plus d’informations sur l’utilisation des API JavaScript Excel et des fonctions Excel personnalisées dans un runtime partagé.
- Découvrez l’exemple de modèles et de pratiques [Gérer le ruban et l’interface utilisateur du volet des tâches, puis exécuter le code sur un document ouvert](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario) pour afficher un exemple plus complet de l’exécution JavaScript partagée en action.

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble : exécuter le code de votre complément dans un Runtime JavaScript partagé](custom-functions-shared-overview.md)
