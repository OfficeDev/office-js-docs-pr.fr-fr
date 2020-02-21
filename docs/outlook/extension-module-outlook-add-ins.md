---
title: Compléments Outlook d’extension de module
description: Créez des applications qui s’exécutent dans Outlook pour simplifier l’accès des utilisateurs aux outils d’informations professionnelles et de productivité sans quitter Outlook.
ms.date: 06/04/2019
localization_priority: Normal
ms.openlocfilehash: a7c4031061dd5cd4605220c6839481b3a88e53d9
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166207"
---
# <a name="module-extension-outlook-add-ins"></a><span data-ttu-id="edc9b-103">Compléments Outlook d’extension de module</span><span class="sxs-lookup"><span data-stu-id="edc9b-103">Module extension Outlook add-ins</span></span>

<span data-ttu-id="edc9b-104">Les compléments d’extension de module figurent dans la barre de navigation Outlook, en regard des onglets Courrier, Tâches et Calendriers.</span><span class="sxs-lookup"><span data-stu-id="edc9b-104">Module extension add-ins appear in the Outlook navigation bar, right alongside mail, tasks, and calendars.</span></span> <span data-ttu-id="edc9b-105">Une extension de module n’utilise pas seulement les informations de courrier et de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="edc9b-105">A module extension is not limited to using mail and appointment information.</span></span> <span data-ttu-id="edc9b-106">Vous pouvez créer des applications qui s’exécutent dans Outlook pour simplifier l’accès des utilisateurs aux outils d’informations professionnelles et de productivité sans quitter Outlook.</span><span class="sxs-lookup"><span data-stu-id="edc9b-106">You can create applications that run inside Outlook to make it easy for your users to access business information and productivity tools without ever leaving Outlook.</span></span>

> [!NOTE]
> <span data-ttu-id="edc9b-107">Les extensions de module sont uniquement prises en charge par Outlook 2016 ou version ultérieure sous Windows.</span><span class="sxs-lookup"><span data-stu-id="edc9b-107">Module extensions are only supported by Outlook 2016 or later on Windows.</span></span>  

## <a name="open-a-module-extension"></a><span data-ttu-id="edc9b-108">Ouvrir une extension de module</span><span class="sxs-lookup"><span data-stu-id="edc9b-108">Open a module extension</span></span>

<span data-ttu-id="edc9b-p102">Pour ouvrir une extension de module, les utilisateurs doivent cliquer sur le nom ou l’icône du module dans la barre de navigation Outlook. Si la navigation compacte est sélectionnée, la barre de navigation affiche une icône indiquant qu’une extension est chargée.</span><span class="sxs-lookup"><span data-stu-id="edc9b-p102">To open a module extension, users click on the module's name or icon in the Outlook navigation bar. If the user has compact navigation selected, the navigation bar has an icon that shows an extension is loaded.</span></span>

![Affiche la barre de navigation compacte lorsqu’une extension de module est chargée dans Outlook.](../images/outlook-module-navigationbar-compact.png)

<span data-ttu-id="edc9b-112">Si l’utilisateur n’utilise pas la navigation compacte, la barre de navigation se présente de deux façons.</span><span class="sxs-lookup"><span data-stu-id="edc9b-112">If the user is not using compact navigation, the navigation bar has two looks.</span></span> <span data-ttu-id="edc9b-113">Si une extension est chargée, elle affiche le nom du complément.</span><span class="sxs-lookup"><span data-stu-id="edc9b-113">With one extension loaded, it shows the name of the add-in.</span></span>

![Affiche la barre de navigation développée lorsqu’une extension de module est chargée dans Outlook.](../images/outlook-module-navigationbar-one.png)

<span data-ttu-id="edc9b-115">Lorsque plusieurs compléments sont chargés, elle affiche le mot **Compléments**. Si vous cliquez sur l’un ou l’autre, l’interface utilisateur de l’extension s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="edc9b-115">When more than one add-in is loaded, it shows the word **Add-ins**. Clicking either will open the extension's user interface.</span></span>

![Affiche la barre de navigation développée lorsque plusieurs extensions de module sont chargées dans Outlook.](../images/outlook-module-navigationbar-more.png)

<span data-ttu-id="edc9b-117">Lorsque vous cliquez sur une extension, Outlook remplace le module intégré par votre module personnalisé pour permettre aux utilisateurs d’interagir avec le complément.</span><span class="sxs-lookup"><span data-stu-id="edc9b-117">When you click on an extension, Outlook replaces the built-in module with your custom module so that your users can interact with the add-in.</span></span> <span data-ttu-id="edc9b-118">Vous pouvez utiliser toutes les fonctionnalités de l’interface API JavaScript pour Outlook dans votre complément et créer des boutons de commande dans le ruban Outlook pour interagir avec le contenu du complément.</span><span class="sxs-lookup"><span data-stu-id="edc9b-118">You can use all of the features of the Outlook JavaScript API in your add-in, and can create command buttons in the Outlook ribbon that will interact with the add-in content.</span></span> <span data-ttu-id="edc9b-119">Les captures d’écran ci-dessous montrent un complément intégré dans la barre de navigation Outlook et comportant des commandes de ruban qui mettent à jour le contenu du complément.</span><span class="sxs-lookup"><span data-stu-id="edc9b-119">The following screenshot shows an add-in that is integrated in the Outlook navigation bar and has ribbon commands that will update the content of the add-in.</span></span>

![Affiche l’interface utilisateur d’une extension de module.](../images/outlook-module-extension.png)

## <a name="example"></a><span data-ttu-id="edc9b-121">Exemple</span><span class="sxs-lookup"><span data-stu-id="edc9b-121">Example</span></span>

<span data-ttu-id="edc9b-122">Vous trouverez ci-dessous une section d’un fichier de manifeste qui définit une extension de module.</span><span class="sxs-lookup"><span data-stu-id="edc9b-122">The following is a section of a manifest file that defines a module extension.</span></span>

```xml
<!-- Add Outlook module extension point -->
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"
                  xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1"
                    xsi:type="VersionOverridesV1_1">

    <!-- Begin override of existing elements -->
    <Description resid="residVersionOverrideDesc" />

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <!-- End override of existing elements -->

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <!-- Set the URL of the file that contains the
                JavaScript function that controls the extension -->
          <FunctionFile resid="residFunctionFileUrl" />

          <!--New Extension Point - Module for a ModuleApp -->
          <ExtensionPoint xsi:type="Module">
            <SourceLocation resid="residExtensionPointUrl" />
            <Label resid="residExtensionPointLabel" />

            <CommandSurface>
              <CustomTab id="idTab">
                <Group id="idGroup">
                  <Label resid="residGroupLabel" />

                  <Control xsi:type="Button" id="group.changeToAssociate">
                    <Label resid="residChangeToAssociateLabel" />
                    <Supertip>
                      <Title resid="residChangeToAssociateLabel" />
                      <Description resid="residChangeToAssociateDesc" />
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="residAssociateIcon16" />
                      <bt:Image size="32" resid="residAssociateIcon32" />
                      <bt:Image size="80" resid="residAssociateIcon80" />
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>changeToAssociateRate</FunctionName>
                    </Action>
                  </Control>
                  
              </Group>
                <Label resid="residCustomTabLabel" />
              </CustomTab>
            </CommandSurface>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="residAddinIcon16" 
                  DefaultValue="https://localhost:8080/Executive-16.png" />
        <bt:Image id="residAddinIcon32" 
                  DefaultValue="https://localhost:8080/Executive-32.png" />
        <bt:Image id="residAddinIcon80" 
                  DefaultValue="https://localhost:8080/Executive-80.png" />
      
        <bt:Image id="residAssociateIcon16" 
                  DefaultValue="https://localhost:8080/Associate-16.png" />
        <bt:Image id="residAssociateIcon32" 
                  DefaultValue="https://localhost:8080/Associate-32.png" />
        <bt:Image id="residAssociateIcon80" 
                  DefaultValue="https://localhost:8080/Associate-80.png" />
      </bt:Images>

      <bt:Urls>
        <bt:Url id="residFunctionFileUrl" 
                DefaultValue="https://localhost:8080/" />
        <bt:Url id="residExtensionPointUrl" 
                DefaultValue="https://localhost:8080/" />
      </bt:Urls>

      <!--Short strings must be less than 30 characters long -->
      <bt:ShortStrings>
        <bt:String id="residExtensionPointLabel" 
                    DefaultValue="Billable Hours" />
        <bt:String id="residGroupLabel" 
                    DefaultValue="Change billing rate" />
        <bt:String id="residCustomTabLabel" 
                    DefaultValue="Billable hours" />

        <bt:String id="residChangeToAssociateLabel" 
                    DefaultValue="Associate" />
      </bt:ShortStrings>

      <bt:LongStrings>
        <bt:String id="residVersionOverrideDesc" 
                    DefaultValue="Version override description" />

        <bt:String id="residChangeToAssociateDesc" 
                    DefaultValue="Change to the associate billing rate: $127/hr" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

## <a name="see-also"></a><span data-ttu-id="edc9b-123">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="edc9b-123">See also</span></span>

- [<span data-ttu-id="edc9b-124">Manifestes de complément Outlook</span><span class="sxs-lookup"><span data-stu-id="edc9b-124">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="edc9b-125">Définir des commandes de complément dans votre manifeste</span><span class="sxs-lookup"><span data-stu-id="edc9b-125">Define add-in commands in your manifest</span></span>](../develop/define-add-in-commands.md)
- [<span data-ttu-id="edc9b-126">Exemple d’heures facturables d’extensions de module Outlook</span><span class="sxs-lookup"><span data-stu-id="edc9b-126">Outlook module extensions Billable hours sample</span></span>](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ModuleExtension)
