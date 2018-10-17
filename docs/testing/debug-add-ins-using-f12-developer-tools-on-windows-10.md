---
title: Débogage des compléments avec les outils de développement F12 sur Windows 10
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 3df245fcd651ec227e0a32d53da186ee332beb8f
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579841"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>Débogage des compléments avec les outils de développement F12 sur Windows 10

Les outils de développement F12 inclus dans Windows 10 vous aident à déboguer, tester et accélérer vos pages web. Ils vous aident également à développer et déboguer les compléments Office si vous n’utilisez pas un IDE comme Visual Studio ou si vous devez examiner un problème pendant l’exécution de votre complément hors de l’IDE.  Dans cet article, vous découvrirez comment utiliser le débogueur des outils de développement F12 de Windows 10 pour tester votre complément Office.

> [!NOTE]
> Les instructions fournies dans cet article ne peuvent pas être utilisées pour déboguer un complément Outlook qui utilise des fonctions d’exécution. Pour déboguer un complément Outlook qui utilise des fonctions d’exécution, nous vous recommandons de vous connecter à Visual Studio en mode script ou à un autre débogueur de script.

## <a name="prerequisites"></a>Conditions préalables

Les logiciels suivants doivent être installés :

- Les outils de développement F12, inclus dans Windows 10. 
    
- L’application cliente Office qui héberge votre complément. 
    
- Votre complément. 

## <a name="using-the-debugger"></a>Utilisation du débogueur

Vous pouvez utiliser le débogueur des outils de développement F12  de Windows 10 pour tester les compléments d’AppSource ou les compléments que vous avez ajoutés à partir d’autres emplacements. Vous pouvez démarrer les outils de développement F12 après l’exécution de votre complément. Les outils F12 s’ouvrent dans une fenêtre séparée et n’utilisent pas Visual Studio.

> [!NOTE]
> Le débogueur fait partie des outils de développement F12 de Windows 10 et d’Internet Explorer. Il n’est pas inclus dans les versions antérieures de Windows. 

Cet exemple utilise Word et un complément gratuit d’AppSource.

1. Ouvrez un document vierge dans Word. 
    
2. Sous l’onglet **Insertion** , dans le groupe Compléments, cliquez sur **Store** et sélectionnez le complément **QR4Office**. (Vous pouvez charger n’importe quel complément depuis le Store ou votre catalogue de compléments.)
    
3. Ouvrez les outils de développement F12 correspondant à votre version d’Office :
    
   - Pour la version 32 bits d'Office, utilisez C:\Windows\System32\F12\IEChooser.exe
    
   - Pour la version 64 bits d'Office, utilisez C:\Windows\SysWOW64\F12\IEChooser.exe
    
   Lorsque vouslancez F12Chooser, une autre fenêtre (intitulée « Choisir la cible à déboguer ») affiche les éventuelles applications à débogue. Sélectionnez l’application qui vous intéresse. Si vous écrivez votre propre complément, sélectionnez le site web où le complément est déployé. Il peut s’agir d’une URL localhost. 
    
   Par exemple, sélectionnez **home.html**. 
    
   ![Écran IEChooser, pointant sur le complément bulles](../images/choose-target-to-debug.png)

4. Dans la fenêtre F12, sélectionnez le fichier à déboguer.
    
   Pour sélectionner le fichier dans la fenêtre F12, cliquez sur l’icône de dossier située au-dessus du volet (gauche) du **script**. Dans la liste des fichiers disponibles affichés dans la liste déroulante, sélectionnez **Home.js**.
    
5. Définissez le point d’arrêt.
    
   Pour définir le point d’arrêt dans **Home.js**, choisissez la ligne 144 située dans la fonction `textChanged`. Vous verrez un point rouge à gauche de la ligne et une ligne correspondante dans le volet **Pile d’appels et Points d’arrêt** (en bas à droite). Pour connaître d’autres façons de définir un point d’arrêt, consultez la rubrique [Inspecter le code JavaScript en cours d’exécution avec le débogueur](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)). 
    
   ![Débogueur avec le point d’arrêt dans le fichier home.js](../images/debugger-home-js-02.png)

6. Exécutez votre complément pour déclencher le point d’arrêt.
    
   Dans Word, cliquez sur la zone de texte URL dans la partie supérieure du volet **QR4Office** et essayez de saisir du texte. Dans le débogueur, dans le volet **Pile d’appels et Points d’arrêt**, vous verrez que le point d’arrêt s’est déclenché et affiche différentes informations. Vous devrez peut-être actualiser le débogueur pour afficher les résultats.
    
   ![Débogueur avec les résultats du point d’arrêt déclenché](../images/debugger-home-js-01.png)


## <a name="see-also"></a>Voir aussi

- [Inspecter le code JavaScript en cours d’exécution avec le débogueur](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [Utilisation des outils de développement F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
