---
title: Octroi du consentement administrateur pour le complément
description: Découvrez comment accorder le consentement de l’administrateur à votre add-in.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 85a230ffd3769b0013081067c88d65d38d43b760
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743782"
---
# <a name="grant-administrator-consent-to-the-add-in"></a>Octroi du consentement administrateur pour le complément

> [!NOTE]
> Cette procédure est uniquement nécessaire quand vous développez le complément. Lorsque votre application de production est déployée sur AppSource ou le Centre d'administration Microsoft 365, les utilisateurs l’utilisent individuellement ou un administrateur consent à l’organisation lors de l’installation.

Effectuez cette procédure *une fois* [que vous avez inscrit le module.](../develop/register-sso-add-in-aad-v2.md)

1. Accédez à la page [Portail Azure - Inscriptions d’applications](https://go.microsoft.com/fwlink/?linkid=2083908) pour afficher l’inscription de votre application.

1. Connectez-vous avec ***les informations d’identification*** d’administrateur à Microsoft 365 location. Par exemple, MonNom@contoso.onmicrosoft.com.

1. Sélectionnez l’application avec le **nom complet $ADD-IN-NAME$**.

1. Sur la page **$ADD-IN-NAME$** , sélectionnez les **autorisations d’API** , puis, sous la section **Autorisations configurées** , choisissez Accorder le consentement administrateur pour [nom du **client].**. **Sélectionnez Oui** pour la confirmation qui s’affiche.

> [!NOTE]
> Nous vous recommandons d’utiliser cette procédure comme meilleure pratique si vous utilisez un compte [Microsoft 365 développeur.](https://developer.microsoft.com/microsoft-365/dev-program) Toutefois, si vous préférez, il est possible de recharger une version de chargement de version de l’ment d’un SSO en cours de développement et d’inviter l’utilisateur avec un formulaire de consentement. Pour plus d’informations, voir [Sideload on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) and [Sideload on Office sur le Web](../testing/sideload-office-add-ins-for-testing.md).
