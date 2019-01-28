---
title: Débogage des compléments avec les outils de développement F12 sur Windows 10
description: ''
ms.date: 10/16/2018
localization_priority: Priority
ms.openlocfilehash: e2378a0449ea33551051b9c3788b84b23a51feb8
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386903"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>Débogage des compléments avec les outils de développement F12 sur Windows 10

Les outils de d?veloppement F12 inclus dans Windows 10 vous aident ? d?boguer, tester et acc?l?rer vos pages web. Ils vous aident ?galement ? d?velopper et d?boguer les compl?ments Office si vous n?utilisez pas un IDE comme Visual Studio ou si vous devez examiner un probl?me pendant l?ex?cution de votre compl?ment hors de l?IDE. Vous pouvez lancer les outils de d?veloppement F12 apr?s l?ex?cution de votre compl?ment.

> [!NOTE]
> Les instructions décrites dans cet article ne peuvent pas être utilisées pour déboguer un complément Outlook qui utilise des fonctions Exécuter. Pour déboguer un complément Outlook qui utilise des fonctions Exécuter, nous vous recommandons de l’attacher à Visual Studio en mode script ou à un autre débogueur de script.

## <a name="prerequisites"></a>Conditions requises

Les logiciels suivants doivent être installés :

- Les outils de développement F12, inclus dans Windows 10. 
    
- L’application cliente Office qui héberge votre complément. 
    
- Votre complément. 

## <a name="using-the-debugger"></a>Utilisation du débogueur

Dans cet article, vous d?couvrirez comment utiliser le d?bogueur des outils de d?veloppement F12 de Windows 10 pour tester votre compl?ment Office. Vous pouvez tester les compl?ments d?AppSource ou des compl?ments que vous avez ajout?s ? partir d?autres emplacements. Les outils F12 s?ouvrent dans une fen?tre s?par?e et n?utilisent pas Visual Studio. Vous pouvez lancer les outils de développement F12 après l’exécution de votre complément. Les outils F12 s’ouvrent dans une fenêtre distincte et n’utilisent pas Visual Studio.

> [!NOTE]
> Le débogueur fait partie des outils de développement F12 de Windows 10 et d’Internet Explorer. Il n’est pas inclus dans les versions antérieures de Windows. 

Cet exemple utilise Word et un complément gratuit d’AppSource.

1. Ouvrez un document vierge dans Word.  
    
2. Sous l’onglet **Insertion**, dans le groupe Compléments, cliquez sur **Store** et sélectionnez le complément **QR4Office**. (Vous pouvez charger n’importe quel complément depuis l’Office Store ou votre catalogue de compléments.)
    
3. Ouvrez les outils de développement F12 correspondant à votre version d’Office :
    
   - Pour la version 32 bits, utilisez C:\Windows\System32\F12\IEChooser.exe
    
   - Pour la version 64 bits, utilisez C:\Windows\SysWOW64\F12\IEChooser.exe
    
   Lorsque vous cliquez sur IEChooser, une autre fenêtre (intitulée « Choisir la cible à déboguer ») affiche les applications possibles pour effectuer le débogage. Sélectionnez l’application de votre choix. Si vous écrivez votre propre complément, sélectionnez le site web où le complément est déployé. Il peut s’agir d’une URL localhost. 
    
   Par exemple, sélectionnez **home.html**. 
    
   ![Écran IEChooser, pointant sur le complément bulles](../images/choose-target-to-debug.png)

4. Dans la fenêtre F12, sélectionnez le fichier à déboguer.
    
   Pour sélectionner le fichier dans la fenêtre F12, cliquez sur l’icône de dossier située au-dessus du volet (gauche) du **script**. Dans la liste des fichiers disponibles qui apparaît dans la liste déroulante, sélectionnez **Home.js**.
    
5. Définissez le point d’arrêt.
    
   Pour définir le point d’arrêt dans **Home.js**, choisissez la ligne 144 située dans la fonction `textChanged`. Vous verrez un point rouge à gauche de la ligne et une ligne correspondante dans le volet Pile d’appels et Points d’arrêt (en bas à droite). Pour connaître d’autres façons de définir un point d’arrêt, consultez la rubrique [Inspecter le code JavaScript en cours d’exécution avec le débogueur](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)). 
    
   ![Débogueur avec le point d’arrêt dans le fichier home.js](../images/debugger-home-js-02.png)

6. Exécutez votre complément pour déclencher le point d’arrêt.
    
   Dans Word, cliquez sur la zone de texte URL dans la partie supérieure du volet **QR4Office** et essayez de saisir du texte. Dans le débogueur, dans le volet **Pile d’appels et Points d’arrêt**, vous verrez que le point d’arrêt s’est déclenché et affiche différentes informations. Vous devrez peut-être actualiser le débogueur pour afficher les résultats.
    
   ![Débogueur avec les résultats du point d’arrêt déclenché](../images/debugger-home-js-01.png)


## <a name="see-also"></a>Voir aussi

- [Inspecter le code JavaScript en cours d’exécution avec le débogueur](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [Utilisation des outils de développement F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
