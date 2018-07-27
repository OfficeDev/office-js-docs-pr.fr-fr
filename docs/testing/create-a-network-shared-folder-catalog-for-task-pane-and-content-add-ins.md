---
title: Chargement de compléments Office pour des tests
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: e5769ef40868ec996194725d98913e61b76279bc
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/27/2018
ms.locfileid: "21270292"
---
# <a name="sideload-office-add-ins-for-testing"></a>Chargez les compléments Office en version test effectuer des tests

Vous pouvez installer un complément Office pour tester dans un client Office s'exécutant sous Windows par l'une des méthodes suivantes :

- Utilisez un catalogue de dossiers partagés pour publier le manifeste sur un partage de fichiers réseau (instructions ci-dessous)
- [Exécutez la commande **« npm run sideload »** à partir de la racine du dossier de projet du complément.](sideload-office-addin-using-sideload-command.md) 
>[!NOTE]
>La méthode « npm run sideload » ne fonctionne que pour les compléments Excel, Word et PowerPoint).

Si vous ne testez pas un complément Word, Excel ou PowerPoint sous Windows, consultez une des rubriques suivantes pour charger la version test de votre complément :

- [Chargement de version test des compléments Office dans Office Online](sideload-office-add-ins-for-testing.md)
- [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)

La vidéo suivante présente la procédure de chargement indépendant de votre complément dans la version de bureau Office ou Office Online à l'aide du catalogue d'un dossier partagé.  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a>Partager un dossier

1. Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.

2. Ouvrez le menu contextuel du dossier (clic droit), puis choisissez **Propriétés**.

3. Ouvrez l’onglet **Partage**.

4. Dans la page **Choisir les utilisateurs...**, ajoutez votre nom et celui des utilisateurs avec lesquels vous souhaitez partager votre complément. S’ils sont tous membres d’un groupe de sécurité, vous pouvez ajouter le groupe. Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier. 

5. Choisissez **Partager** > **Terminer** > **Fermer**.


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>Spécifier le dossier partagé en tant que catalogue approuvé
      
1. Ouvrez un nouveau document dans Excel, Word ou PowerPoint.
    
2. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
    
3. Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.
    
4. Choisissez **Catalogues de compléments approuvés**.
    
5. Dans la zone **URL du catalogue**, entrez le chemin d’accès réseau complet au catalogue de dossiers partagés, puis choisissez **Ajouter un catalogue**.
    
6. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.

7. Fermez l’application Office afin que vos modifications prennent effet.
    

## <a name="sideload-your-add-in"></a>Charger votre complément

1. Placez le fichier manifeste d’un complément en cours de test dans le catalogue de dossiers partagés. Notez que vous déployez l’application web sur un serveur web. Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban.

3. Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.

4. Sélectionnez le nom du complément, puis choisissez **OK** pour insérer le complément.


## <a name="see-also"></a>Voir aussi

- [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md)
- [Publier votre complément Office](../publish/publish.md)
    
