---
title: D?bogage des compl?ments avec les outils de d?veloppement F12 sur Windows 10
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e1e4cde4a1a0fe27058346b93e8aaa39dd75a4e3
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>D?bogage des compl?ments avec les outils de d?veloppement F12 sur Windows 10

Les outils de d?veloppement F12 inclus dans Windows 10 vous aident ? d?boguer, tester et acc?l?rer vos pages web. Ils vous aident ?galement ? d?velopper et d?boguer les compl?ments Office si vous n?utilisez pas un IDE comme Visual Studio ou si vous devez examiner un probl?me pendant l?ex?cution de votre compl?ment hors de l?IDE. Vous pouvez lancer les outils de d?veloppement F12 apr?s l?ex?cution de votre compl?ment.

Dans cet article, vous d?couvrirez comment utiliser le d?bogueur des outils de d?veloppement F12 de Windows 10 pour tester votre compl?ment Office. Vous pouvez tester les compl?ments d?AppSource ou des compl?ments que vous avez ajout?s ? partir d?autres emplacements. Les outils F12 s?ouvrent dans une fen?tre s?par?e et n?utilisent pas Visual Studio.

> [!NOTE]
> Le d?bogueur fait partie des outils de d?veloppement F12 de Windows 10 et d?Internet Explorer. Il n?est pas inclus dans les versions ant?rieures de Windows. 

## <a name="prerequisites"></a>Conditions pr?alables

Les logiciels suivants doivent ?tre install?s :

- Les outils de d?veloppement F12, inclus dans Windows 10. 
    
- L?application cliente Office qui h?berge votre compl?ment. 
    
- Votre compl?ment. 

## <a name="using-the-debugger"></a>Utilisation du d?bogueur

Cet exemple utilise Word et un compl?ment gratuit d?AppSource.

1. Ouvrez un document vierge dans Word. 
    
2. Sous l?onglet **Insertion**, dans le groupe Compl?ments, cliquez sur **Store** et s?lectionnez le compl?ment QR4Office. (Vous pouvez charger n?importe quel compl?ment depuis le Store ou votre catalogue de compl?ments.)
    
3. Ouvrez les outils de d?veloppement F12 correspondant ? votre version d?Office :
    
   - Pour la version 32 bits, utilisez C:\Windows\System32\F12\F12Chooser.exe
    
   - Pour la version 64 bits, utilisez C:\Windows\SysWOW64\F12\F12Chooser.exe
    
   Lorsque vous cliquez sur F12Chooser, une autre fen?tre (intitul?e ? Choisir la cible ? d?boguer ?) affiche les applications possibles pour effectuer le d?bogage. S?lectionnez l?application de votre choix. Si vous ?crivez votre propre compl?ment, s?lectionnez le site web o? le compl?ment est d?ploy?. Il peut s?agir d?une URL localhost. 
    
   Par exemple, s?lectionnez **home.html**. 
    
   ![?cran du s?lecteur F12, pointe vers un compl?ment de type ? bulles ?](../images/choose-target-to-debug.png)

4. Dans la fen?tre F12, s?lectionnez le fichier ? d?boguer.
    
   Pour s?lectionner le fichier, cliquez sur l?ic?ne de dossier situ?e au-dessus du volet (gauche) du **script**. La liste d?roulante affiche les fichiers disponibles. S?lectionnez home.js.
    
5. D?finissez le point d?arr?t.
    
   Pour d?finir un point d'arr?t dans home.js, choisissez la ligne 144 qui se trouve dans la fonction _textChanged_. Vous verrez un point rouge ? gauche de la ligne et une ligne correspondante dans le volet **Callstack and Breakpoints** (en bas ? droite). Pour conna?tre d'autres mani?res de d?finir un point d'arr?t, r?f?rez-vous ? [Consulter JavaScript en fonctionnement avec le d?bogueur](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx). 
    
   ![D?bogueur avec le point d?arr?t dans le fichier home.js](../images/debugger-home-js-02.png)

6. Ex?cutez votre compl?ment pour d?clencher le point d?arr?t.
    
   Cliquez sur la zone de texte URL dans la partie sup?rieure du volet QR4Office pour modifier le texte. Dans le d?bogueur, dans le volet **Pile d?appels et Points d?arr?t**, vous verrez que le point d?arr?t s?est d?clench? et affiche diff?rentes informations. Vous devrez peut-?tre actualiser l?outil F12 pour afficher les r?sultats.
    
   ![D?bogueur avec les r?sultats du point d?arr?t d?clench?](../images/debugger-home-js-01.png)


## <a name="see-also"></a>Voir aussi

- [Inspecter le code JavaScript en cours d?ex?cution avec le d?bogueur](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx)
- [Utilisation des outils de d?veloppement F12](https://msdn.microsoft.com/en-us/library/bg182326%28v=vs.85%29.aspx)
    
