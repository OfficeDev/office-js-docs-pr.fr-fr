---
title: Débogage des compléments avec les outils de développement F12 sur Windows 10
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e1e4cde4a1a0fe27058346b93e8aaa39dd75a4e3
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19438724"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>Débogage des compléments avec les outils de développement F12 sur Windows 10

Les outils de développement F12 inclus dans Windows 10 vous aident à déboguer, tester et accélérer vos pages web. Ils vous aident également à développer et déboguer les compléments Office si vous n’utilisez pas un IDE comme Visual Studio ou si vous devez examiner un problème pendant l’exécution de votre complément hors de l’IDE. Vous pouvez lancer les outils de développement F12 après l’exécution de votre complément.

Dans cet article, vous découvrirez comment utiliser le débogueur des outils de développement F12 de Windows 10 pour tester votre complément Office. Vous pouvez tester les compléments d’AppSource ou des compléments que vous avez ajoutés à partir d’autres emplacements. Les outils F12 s’ouvrent dans une fenêtre séparée et n’utilisent pas Visual Studio.

> [!NOTE]
> Le débogueur fait partie des outils de développement F12 de Windows 10 et d’Internet Explorer. Il n’est pas inclus dans les versions antérieures de Windows. 

## <a name="prerequisites"></a>Conditions préalables

Les logiciels suivants doivent être installés :

- Les outils de développement F12, inclus dans Windows 10. 
    
- L’application cliente Office qui héberge votre complément. 
    
- Votre complément. 

## <a name="using-the-debugger"></a>Utilisation du débogueur

Cet exemple utilise Word et un complément gratuit d’AppSource.

1. Ouvrez un document vierge dans Word. 
    
2. Sous l’onglet **Insertion**, dans le groupe Compléments, cliquez sur **Store** et sélectionnez le complément QR4Office. (Vous pouvez charger n’importe quel complément depuis le Store ou votre catalogue de compléments.)
    
3. Ouvrez les outils de développement F12 correspondant à votre version d’Office :
    
   - Pour la version 32 bits, utilisez C:\Windows\System32\F12\F12Chooser.exe
    
   - Pour la version 64 bits, utilisez C:\Windows\SysWOW64\F12\F12Chooser.exe
    
   Lorsque vous cliquez sur F12Chooser, une autre fenêtre (intitulée « Choisir la cible à déboguer ») affiche les applications possibles pour effectuer le débogage. Sélectionnez l’application de votre choix. Si vous écrivez votre propre complément, sélectionnez le site web où le complément est déployé. Il peut s’agir d’une URL localhost. 
    
   Par exemple, sélectionnez **home.html**. 
    
   ![Écran du sélecteur F12, pointe vers un complément de type « bulles »](../images/choose-target-to-debug.png)

4. Dans la fenêtre F12, sélectionnez le fichier à déboguer.
    
   Pour sélectionner le fichier, cliquez sur l’icône de dossier située au-dessus du volet (gauche) du **script**. La liste déroulante affiche les fichiers disponibles. Sélectionnez home.js.
    
5. Définissez le point d’arrêt.
    
   Pour définir un point d'arrêt dans home.js, choisissez la ligne 144 qui se trouve dans la fonction _textChanged_. Vous verrez un point rouge à gauche de la ligne et une ligne correspondante dans le volet **Callstack and Breakpoints** (en bas à droite). Pour connaître d'autres manières de définir un point d'arrêt, référez-vous à [Consulter JavaScript en fonctionnement avec le débogueur](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx). 
    
   ![Débogueur avec le point d’arrêt dans le fichier home.js](../images/debugger-home-js-02.png)

6. Exécutez votre complément pour déclencher le point d’arrêt.
    
   Cliquez sur la zone de texte URL dans la partie supérieure du volet QR4Office pour modifier le texte. Dans le débogueur, dans le volet **Pile d’appels et Points d’arrêt**, vous verrez que le point d’arrêt s’est déclenché et affiche différentes informations. Vous devrez peut-être actualiser l’outil F12 pour afficher les résultats.
    
   ![Débogueur avec les résultats du point d’arrêt déclenché](../images/debugger-home-js-01.png)


## <a name="see-also"></a>Voir aussi

- [Inspecter le code JavaScript en cours d’exécution avec le débogueur](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx)
- [Utilisation des outils de développement F12](https://msdn.microsoft.com/en-us/library/bg182326%28v=vs.85%29.aspx)
    
