---
title: Chargement de version test des compl?ments Office dans Office Online
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 823821f990674a2d822508a860a7e5d6424e0245
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a>Chargement de version test des compl?ments Office dans Office Online

Vous pouvez installer un compl?ment Office test sans avoir ? le placer au pr?alable dans un catalogue de compl?ments en utilisant le chargement de version test. Le chargement de version test peut ?tre effectu? sur Office 365 ou Office Online. La proc?dure pr?sente de l?g?res diff?rences d?une plateforme ? l?autre. 

Lorsque vous chargez une version test d?un compl?ment, le manifeste du compl?ment est stock? dans le stockage local du navigateur. Ainsi, si vous videz le cache du navigateur ou si vous basculez vers un autre navigateur, vous devez ? nouveau charger une version test de compl?ment.


> [!NOTE]
> Tel que d?crit dans cet article, le chargement de version test est pris en charge dans Word, Excel et PowerPoint. Pour charger une version test de compl?ment Outlook, voir la rubrique relative au [chargement de version test des compl?ments Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing).

La vid?o suivante pr?sente la proc?dure de chargement de version test de votre compl?ment dans la version de bureau Office ou Office Online.  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-on-office-365"></a>Chargement de version test d?un compl?ment Office dans Office 365


1. Connectez-vous ? votre compte Office 365.
    
2. Ouvrez le lanceur d?applications ? l?extr?mit? gauche de la barre d?outils et s?lectionnez **Excel**,  **Word** ou **PowerPoint**, puis cr?ez un document.
    
3. Ouvrez l?onglet **Ins?rer** dans le ruban, puis dans la section **Compl?ments**, choisissez **Compl?ments Office**.
    
4. Dans la bo?te de dialogue **Compl?ments Office**, s?lectionnez l?onglet **MON ORGANISATION**, puis **T?l?charger mon compl?ment**.
    
    ![Bo?te de dialogue intitul?e Compl?ment Office avec un lien dans le coin sup?rieur gauche indiquant ? Charger mon compl?ment ?.](../images/office-add-ins.png)

5.  **Acc?dez** au fichier manifeste du compl?ment, puis s?lectionnez **T?l?charger**.
    
    ![Bo?te de dialogue de chargement de compl?ment avec des boutons pour parcourir, t?l?charger et annuler.](../images/upload-add-in.png)

6. Verify that your compl?ment is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.
    

## <a name="sideload-an-office-add-in-on-office-online"></a>Charger une version test d?un compl?ment Office sur Office Online


1. Ouvrez [Microsoft Office Online](https://office.live.com/).
    
2. Dans **Commencer ? utiliser les applications en ligne maintenant**, choisissez **Excel**, **Word** ou **PowerPoint**, puis ouvrez un document.
    
3. Ouvrez l?onglet **Ins?rer** dans le ruban, puis dans la section **Compl?ments**, choisissez **Compl?ments Office**.
    
4. Dans la bo?te de dialogue **Compl?ments Office**, s?lectionnez l?onglet **MES COMPL?MENTS**, choisissez **G?rer mes compl?ments**, puis **T?l?charger mon compl?ment**.
    
    ![Bo?te de dialogue Compl?ments Office avec une liste d?roulante dans le coin sup?rieur droit indiquant ? G?rer mes compl?ments ? et une autre liste d?roulante sous cette derni?re avec l?option ? Charger mon compl?ment ?](../images/office-add-ins-my-account.png)

5.  **Acc?dez** au fichier manifeste du compl?ment, puis s?lectionnez **T?l?charger**.
    
    ![Bo?te de dialogue de t?l?chargement de compl?ment avec des boutons pour parcourir, t?l?charger et annuler.](../images/upload-add-in.png)

6. V?rifiez que votre compl?ment est install?. S?il s?agit d?une commande de compl?ment, elle doit appara?tre dans le ruban ou dans le menu contextuel. S?il s?agit d?un compl?ment du volet Office, le volet doit appara?tre.

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Chargement d?une version test d?un compl?ment lors de l?utilisation de Visual Studio

Si vous d?veloppez votre compl?ment ? l?aide de Visual Studio, le processus de chargement d?une version de teste est similaire. La seule diff?rence est que vous devez mettre ? jour la valeur de l??l?ment **SourceURL** dans votre manifeste, de sorte ? inclure l?URL enti?re de l?emplacement de d?ploiement du compl?ment. 

Si vous ?tes en train de d?velopper votre compl?ment, recherchez-le dans le fichier manifest.xml et mettez ? jour la valeur de l??l?ment **SourceLocation** de fa?on ? inclure un URI absolu. Visual Studio met en place un jeton pour votre d?ploiement localhost.

Par exemple : 

```xml
<SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
```
