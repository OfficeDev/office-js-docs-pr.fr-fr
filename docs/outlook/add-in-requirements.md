---
title: Conditions requises pour les compléments Outlook
description: Pour que les compléments Outlook se chargent et fonctionnent correctement, les serveurs et les clients doivent répondre à plusieurs conditions.
ms.date: 10/09/2019
localization_priority: Priority
ms.openlocfilehash: be93ef69e60530947c18b5b5be294c6d12ed06f1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720875"
---
# <a name="outlook-add-in-requirements"></a>Conditions requises pour les compléments Outlook

Pour que les compléments Outlook se chargent et fonctionnent correctement, les serveurs et les clients doivent répondre à plusieurs conditions.

## <a name="client-requirements"></a>Configuration requise du client

- Le client doit être l’un des hôtes pris en charge pour les compléments Outlook. Les clients suivants prennent en charge les compléments :

   - Outlook 2013 ou version ultérieure sur Windows
   - Outlook 2016 ou version ultérieure sur Mac
   - Outlook sur iOS
   - Outlook sur Android
   - Outlook sur le web pour Exchange 2016 ou une version ultérieure et Office 365
   - Outlook sur le web pour Exchange 2013
   - Outlook.com

- Le client doit être connecté à un serveur Exchange ou Office 365 via une connexion directe. Lors de la configuration du client, l’utilisateur doit sélectionner un compte de type **Exchange**, **Office 365** ou **Outlook.com**. Si le client est configuré pour se connecter avec POP3 ou IMAP, les compléments ne seront pas chargés.

## <a name="mail-server-requirements"></a>Configuration requise pour le serveur de messagerie

Si l’utilisateur est connecté à Office 365 ou à Outlook.com, le serveur de messagerie a déjà la configuration requise. En revanche, pour les utilisateurs connectés à des installations locales de Microsoft Exchange Server, les conditions suivantes s’appliquent.

- Le serveur doit être un serveur Exchange 2013 ou de version ultérieure.
- Les services web Exchange doivent être activés et exposés sur Internet. De nombreux compléments exigent que les services web Exchange fonctionnent correctement.
- Le serveur doit avoir un certificat d’authentification valide pour émettre des jetons d’identité valides. Les nouvelles installations du serveur Exchange incluent un certificat d’authentification par défaut. Pour plus d’informations, reportez-vous aux articles [Certificats numériques et chiffrement dans Exchange 2016](/Exchange/architecture/client-access/certificates) et [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).
- Pour accéder à des compléments à partir d’[AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), les serveurs d’accès au client doivent être en mesure de communiquer avec AppSource.

## <a name="add-in-server-requirements"></a>Conditions de serveur pour le complément

Les fichiers du complément (HTML, JavaScript, etc.) peuvent être hébergés sur n’importe quelle plateforme de serveur web. Les seules conditions sont que le serveur doit être configuré de manière à utiliser le protocole HTTPS et que le certificat SSL doit être approuvé par le client.

## <a name="see-also"></a>Voir aussi

- [Configuration requise pour exécuter des compléments Office](../concepts/requirements-for-running-office-add-ins.md)
- [Disponibilité des compléments Office sur les plateformes et les hôtes (section Outlook)](../overview/office-add-in-availability.md#outlook)
- [Prise en charge des ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
