---
title: Cycle de vie du développement des compléments Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 5b056527deaf03beb51d755b582be715fbd14233
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505894"
---
# <a name="office-add-ins-development-lifecycle"></a>Cycle de vie du développement des compléments Office

> [!NOTE]
> Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de validation AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)). 

Le cycle de vie de développement classique d’un complément Office comprend les étapes suivantes :


## <a name="1-decide-on-the-purpose-of-the-add-in"></a>1. Déterminer l’objet du complément
    
Posez-vous les questions suivantes :
    
- Quelle peut être l’utilité du complément ? 
        
- Comment peut-elle contribuer à accroître la productivité de vos clients ?
        
- Quels scénarios sont pris en charge par les fonctionnalités de votre complément ?
    
Déterminez les fonctionnalités et les scénarios les plus importants et réalisez la conception à partir de ces éléments. 

    
## <a name="2-identify-the-data-and-data-source-for-the-add-in"></a>2. Identifier les données et la source de données du complément
    
- Les données figurent-elles dans un document, un classeur, une présentation, un projet ou une base de données Access basée sur un navigateur ? 
    
- Les données concernent-elles un ou plusieurs éléments d’une boîte aux lettres Exchange Server ou Exchange Online ? 
    
- Les données proviennent-elles d’une source externe telle qu’un service web ?

    
## <a name="3-identify-the-type-of-add-in-and-office-host-applications-that-best-support-the-purpose-of-the-add-in"></a>3. Identifier le type de complément et les applications hôtes Office les mieux adaptés pour prendre en charge l’objet de l’application
    
Tenez compte des informations suivantes pour identifier les scénarios :
    
- Les clients utiliseront-ils le complément pour enrichir le contenu d’un document ou d’une base de données Access reposant sur navigateur ? Si c’est le cas, vous pouvez envisager de créer un **complément de contenu**. 
    
- Les clients utiliseront-ils le complément lors de la visualisation ou de la composition d’un message électronique ou d’un rendez-vous ? Est-il important de pouvoir exposer le complément conformément au contexte actuel ? La possibilité de rendre le complément disponible non seulement sur le bureau mais également sur des tablettes ou des smartphones constitue-telle une priorité ?
    
    Si vous répondez oui à l’une de ces questions, envisagez de créer un **complément Outlook**. Identifiez le contexte qui déclenchera votre complément (par exemple, un formulaire de composition utilisé par un utilisateur, des types de messages spécifiques, la présence d’une pièce jointe, l’adresse, la suggestion de tâche, la suggestion de réunion ou certains modèles de chaînes dans le contenu d’un courrier électronique ou d’un rendez-vous). 
        
    Reportez-vous à l’article relatif aux [règles d’activation pour les compléments Outlook](https://docs.microsoft.com/outlook/add-ins/activation-rules) pour savoir comment activer le complément Outlook en fonction du contexte. 
    
- Les clients utiliseront-ils le complément pour améliorer l’affichage ou l’expérience de création d’un document ? Si c’est le cas, vous pouvez créer un **complément de volet Office**. 

La prise en charge pour certaines API de complément peut être différente entre les applications Office et la plateforme d’exécution (Windows, Mac, Web, Mobile). Pour afficher la couverture API actuelle par le client et la plateforme, consultez la page concernant la [disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md).  

    
## <a name="4-design-and-implement-the-user-experience-and-user-interface-for-the-add-in"></a>4. Concevoir et implémenter l’expérience utilisateur et l’interface utilisateur pour le complément
    
Concevez une expérience utilisateur rapide et fluide qui est cohérente, facile à apprendre, avec des scénarios nécessitant uniquement quelques étapes d’exécution. Selon l’objet du complément, utilisez des API ou des services web de tiers.
    
Vous pouvez faire votre choix parmi divers outils de développement web et utiliser du code HTML et JavaScript pour implémenter l’interface utilisateur.

    
## <a name="5-create-an-xml-manifest-file-based-on-the-office-add-ins-manifest-schema"></a>5. Créer un fichier manifeste XML basé sur le schéma de manifeste des compléments Office
    
Créez un manifeste XML pour identifier le complément et sa configuration requise, spécifiez les emplacements du code HTML et des fichiers JavaScript et CSS utilisés par le complément, et précisez la taille par défaut et les autorisations  en fonction du type de complément.
    
Pour les compléments Outlook, vous pouvez spécifier le contexte, en fonction du message ou du rendez-vous actif, sous lequel votre complément est pertinent et doit être disponible dans l’interface utilisateur d’Outlook. Vous devez également choisir les périphériques que votre complément doit prendre en charge. Dans le manifeste, spécifiez le contexte sous forme de règles d’activation, ainsi que les périphériques pris en charge.
    

## <a name="6-install-and-test-the-add-in"></a>6. Installer et tester le complément
    
Placez les fichiers HTML et les éventuels fichiers JavaScript et CSS sur les serveurs web qui sont spécifiés dans le fichier manifeste du complément. Le processus d’installation d’un complément dépend du type de celui-ci. Pour plus d’informations, reportez-vous à la page relative au [chargement d’une version test des compléments Office à des fins de test](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
    
Pour les compléments Outlook, installez-le dans une boîte aux lettres Exchange et spécifiez l’emplacement du fichier manifeste du complément dans le Centre d’administration Exchange (CAE). Pour plus d’informations, consultez la rubrique [Déployer et installer des compléments Outlook à des fins de test](https://docs.microsoft.com/outlook/add-ins/testing-and-tips).

    
## <a name="7-publish-the-add-in"></a>7. Publier le complément
    
Vous pouvez envoyer le complément à AppSource, à partir duquel les clients peuvent installer le complément. En outre, vous pouvez publier le volet Office et les compléments du contenu dans le catalogue de compléments d’un dossier privé dans SharePoint ou dans un dossier réseau partagé, et vous pouvez déployer un complément Outlook directement sur un serveur Exchange pour votre organisation. Pour plus d’informations, consultez la rubrique [Publier votre complément Office](../publish/publish.md).
    
    
## <a name="8-maintain-the-add-in"></a>8. Mettre à jour le complément
    
Si votre complément appelle un service web, et si vous effectuez des mises à jour pour le service web après la publication du complément, vous n’êtes pas obligé de republier le complément. Toutefois, si vous modifiez des éléments ou des données que vous avez soumise pour votre complément, tels que le manifeste de l’application add-in, captures d’écran, icônes, fichiers HTML ou JavaScript, vous devez republier le complément. 
    
En particulier si vous avez publié le complément sur AppSource, vous devez soumettre à nouveau votre complément afin qu’AppSource puisse implémenter les modifications. Vous devez soumettre à nouveau votre complément avec un manifeste de complément mis à jour qui inclut un nouveau numéro de version. Vous devez également veiller à mettre à jour le numéro de version du complément dans le formulaire d’envoi afin qu’il corresponde au numéro de version du nouveau manifeste. Pour les compléments Outlook, vous devez vous assurer que l’élément [Id](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id?view=office-js) contient un autre UUID dans le manifeste de complément.
    
