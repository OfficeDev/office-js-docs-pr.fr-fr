---
title: Manifeste Teams pour les compléments Office (préversion)
description: Obtenez une vue d’ensemble du manifeste JSON en préversion.
ms.date: 06/15/2022
ms.localizationpriority: high
ms.openlocfilehash: 9eb2a886ed700bee0d7ba91d8a2c48e5de92fea1
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092881"
---
# <a name="teams-manifest-for-office-add-ins-preview"></a>Manifeste Teams pour les compléments Office (préversion)

Microsoft apporte un certain nombre d’améliorations à la plateforme de développement Microsoft 365. Ces améliorations offrent une cohérence accrue dans le développement, le déploiement, l’installation et l’administration de tous les types d’extensions de Microsoft 365, y compris les compléments Office. Ces modifications sont compatibles avec les compléments existants. 

Une amélioration importante sur laquelle nous travaillons est la possibilité de créer une unité de distribution unique pour toutes vos extensions Microsoft 365 à l’aide du même format de manifeste et du même schéma, en fonction du manifeste Teams au format JSON actuel.

Nous avons effectué une première étape importante vers ces objectifs en vous permettant de créer des compléments Outlook, s’exécutant uniquement sur Windows, avec une version du manifeste JSON Teams.

> [!NOTE]
> Le nouveau manifeste est disponible en préversion et peut être modifié en fonction des commentaires. Nous encourageons les développeurs de compléments expérimentés à l’expérimenter. Le manifeste d’aperçu ne doit pas être utilisé dans les compléments de production. 

Pendant la période de préversion anticipée, les limitations suivantes s’appliquent.

- La préversion du manifeste Teams prend uniquement en charge les compléments Outlook et uniquement sur l’abonnement Office pour Windows. Nous travaillons à l’extension de la prise en charge à Excel, PowerPoint et Word.
- Il n’est pas encore possible de combiner et de charger une version test d’un complément avec une application Teams, telle qu’un onglet personnel Teams ou d’autres types d’extensions Microsoft 365. Dans les prochains mois, nous continuerons à étendre la préversion pour prendre en charge ces scénarios et fournir des outils supplémentaires pour mettre à jour les manifestes au format d’aperçu.

> [!TIP]
> Vous êtes prêt à commencer à utiliser le manifeste Teams en préversion ? Commencez par [Créer un complément Outlook avec un manifeste Teams (préversion).](../quickstarts/outlook-quickstart-json-manifest.md)

## <a name="overview-of-the-json-manifest"></a>Vue d’ensemble du manifeste JSON

### <a name="schemas-and-general-points"></a>Schémas et points généraux

Il n’existe qu’un seul schéma pour le [manifeste JSON en préversion](/microsoftteams/platform/resources/dev-preview/developer-preview-intro), contrairement au manifeste XML actuel qui a un total de sept [Schémas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).  

### <a name="conceptual-mapping-of-the-preview-json-and-current-xml-manifests"></a>Mappage conceptuel des manifestes JSON et XML actuels en préversion

Cette section décrit le manifeste JSON en préversion pour les lecteurs qui connaissent le manifeste XML actuel. Voici quelques points à garder à l’esprit : 

- JSON ne fait pas la distinction entre l’attribut et la valeur d’élément comme le fait XML. En règle générale, le JSON qui mappe à un élément XML fait de la valeur de l’élément et de chacun des attributs une propriété enfant. L’exemple suivant montre un balisage XML et son équivalent JSON.
  
  ```xml
  <MyThing color="blue">Some text</MyThing>
  ```

  ```json
  "myThing" : {
      "color": "blue",
      "text": "Some text"
  }
  ```
- Il existe de nombreux endroits dans le manifeste XML actuel où un élément avec un nom au pluriel a des enfants avec la version unique du même nom. Par exemple, le balisage pour configurer un menu personnalisé comprend un **\<Items\>** élément qui peut avoir plusieurs **\<Item\>** enfants.. L’équivalent JSON de ces divers éléments est une propriété avec un tableau comme valeur. Les membres du tableau sont des objets *anonymes* , et non des propriétés nommées « item » ou « item1 », « item2 », etc. Voici un exemple.

  ```json
  "items": [
      {
          -- markup for a menu item is here --
      },
      {
          -- markup for another menu item is here --
      }
  ]
  ```

#### <a name="top-level-structure"></a>Structure de niveau supérieur

Le niveau racine du manifeste JSON de l'aperçu, qui correspond approximativement à **\<OfficeApp\>** l'élément du manifeste XML actuel, est un objet anonyme. 

Les enfants **\<OfficeApp\>** de sont communément divisés en deux catégories notionnelles. L'**\<VersionOverrides\>** élément est une catégorie. L'autre est constitué de tous les autres enfants **\<OfficeApp\>** de , qui sont collectivement appelés le manifeste de base. Par conséquent, le manifeste JSON en préversion a une division similaire. Il existe une propriété «extension» de niveau supérieur qui correspond grosso modo à **\<VersionOverrides\>** l'élément dans ses objectifs et ses propriétés enfant. Le manifeste JSON en préversion a également plus de 10 autres propriétés de niveau supérieur qui servent collectivement les mêmes objectifs que le manifeste de base du manifeste XML. Ces autres propriétés peuvent être considérées collectivement comme le manifeste de base du manifeste JSON. 

> [!NOTE]
> Lorsqu’il est possible de combiner un complément avec d’autres types d’extension Microsoft 365 dans un seul manifeste, d’autres propriétés de niveau supérieur ne tiennent pas dans la notion de manifeste de base. Il existe généralement une propriété de niveau supérieur pour chaque type d’extension Microsoft 365, par exemple « configurableTabs », « bots » et « connecteurs ». Pour obtenir des exemples, consultez la [documentation du manifeste Teams](/microsoftteams/platform/resources/schema/manifest-schema). Cette structure indique clairement que la propriété « extension » représente un complément Office en tant que type d’extension Microsoft 365.

#### <a name="base-manifest"></a>Manifeste de base

Les propriétés du manifeste de base spécifient les caractéristiques du complément que *tout* type d’extension de Microsoft 365 doit avoir. Cela inclut les onglets Teams et les extensions de message, pas seulement les compléments Office. Ces caractéristiques incluent un nom public et un ID unique. Le tableau suivant montre un mappage de certaines propriétés de niveau supérieur critiques dans le manifeste JSON en préversion aux éléments XML du manifeste actuel, où le principe de mappage est *l’objectif* du balisage.

|Propriété JSON|Objectif|Éléments XML|Commentaires|
|:-----|:-----|:-----|:-----|
|« $schema »| Identifie le schéma de manifeste. | attributs de **\<OfficeApp\>** et **\<VersionOverrides\>** |*Aucun.* |
|"id"| GUID du complément. | **\<Id\>**|*Aucun.* |
|« version »| Version du complément. | **\<Version\>** |*Aucun.* |
|« manifestVersion »| Version du schéma de manifeste. |  attributs de **\<OfficeApp\>** |*Aucun.* |
|« nom »| Nom public du complément. | **\<DisplayName\>** |*Aucun.* |
|« description »| Description publique du complément.  | **\<Description\>** |*Aucun.* |
|« accentColor »|*Aucun.* |*Aucun.* | Cette propriété n’a pas d’équivalent dans le manifeste XML actuel et n’est pas utilisée dans la préversion du manifeste JSON. Mais il doit être présent. |
|« développeur »| Identifie le développeur du complément. | **\<ProviderName\>** |*Aucun.* |
|« localizationInfo »| Configure les paramètres régionaux par défaut et les autres paramètres régionaux pris en charge. | **\<DefaultLocale\>** et **\<Override\>** |*Aucun.* |
|« webApplicationInfo »| Identifie l’application web du complément telle qu’elle est connue dans Azure Active Directory. | **\<WebApplicationInfo\>** | Dans le manifeste XML actuel, **\<WebApplicationInfo\>** l'élément se trouve à l'intérieur, et **\<VersionOverrides\>** non dans le manifeste de base. |
|« autorisation »| Identifie les autorisations Microsoft Graph dont le complément a besoin. | **\<WebApplicationInfo\>** | Dans le manifeste XML actuel, **\<WebApplicationInfo\>** l'élément se trouve à l'intérieur, et **\<VersionOverrides\>** non dans le manifeste de base. |

Les éléments **\<Hosts\>**, , **\<Requirements\>** et **\<ExtendedOverrides\>** font partie du manifeste de base dans le manifeste XML actuel. Toutefois, les concepts et les objectifs associés à ces éléments sont configurés dans la propriété « extension » du manifeste JSON en préversion.

#### <a name="extension-property"></a>Propriété « extension »

La propriété « extension » dans le manifeste JSON en préversion représente principalement des caractéristiques du complément qui ne seraient pas pertinentes pour d’autres types d’extensions Microsoft 365. Par exemple, les applications Office que le complément étend (par exemple, Excel, PowerPoint, Word et Outlook) sont spécifiées dans la propriété « extension », tout comme les personnalisations du ruban d’application Office. Les objectifs de configuration de la propriété «extension» correspondent étroitement à ceux de **\<VersionOverrides\>** l'élément dans le manifeste XML actuel.

> [!NOTE]
> La **\<VersionOverrides\>** section du manifeste XML actuel comporte un système de «double saut» pour de nombreuses ressources de type chaîne. Les chaînes de caractères, y compris les URL, sont spécifiées et se voient attribuer un ID dans **\<Resources\>** le fils de **\<VersionOverrides\>**. Les éléments qui nécessitent une chaîne de caractères ont un `resid`attribut qui correspond à l' ID d'une chaîne **\<Resources\>** de caractères dans l'élément. La propriété « extension » du manifeste JSON en préversion simplifie les choses en définissant des chaînes directement en tant que valeurs de propriété. Il n'y a rien dans le manifeste JSON qui soit équivalent à **\<Resources\>** l'élément.

Le tableau suivant montre un mappage de certaines propriétés enfants de haut niveau de la propriété « extension » dans le manifeste JSON d’aperçu aux éléments XML du manifeste actuel. La notation par points est utilisée pour référencer les propriétés enfants.

|Propriété JSON|Objectif|Éléments XML|Commentaires|
|:-----|:-----|:-----|:-----|
| « requirements.capabilities » | Identifie les ensembles de conditions requises que le complément doit être installable. | **\<Requirements\>** et **\<Sets\>** |*Aucun.* |
| « étendues des conditions » | Identifie les applications Office dans lesquelles le complément peut être installé. | **\<Hosts\>** |*Aucun.* |
| « rubans » | Rubans personnalisés par le complément. | **\<Hosts\>**, **ExtensionPoints**, et divers **\*éléments** FormFactor | La propriété « rubans » est un tableau d’objets anonymes qui fusionnent chacun les objectifs de ces trois éléments. Consultez le [tableau « rubans](#ribbons-table) ».|
| « alternatives » | Spécifie la compatibilité descendante avec un complément COM équivalent, XLL ou les deux. | **\<EquivalentAddins\>** | Consultez [EquivalentAddins - Consultez également](/javascript/api/manifest/equivalentaddins#see-also) pour obtenir des informations générales. |
| « runtimes »  | Configure différents types de compléments qui ont peu ou pas d’interface utilisateur, tels que des compléments de fonction uniquement personnalisés et [commandes de fonction](../design/add-in-commands.md#types-of-add-in-commands). | **\<Runtimes\>**. **\<FunctionFile\>**, et **\<ExtensionPoint\>** (de type CustomFunctions) |*Aucun.* |
| « AutoRunEvents » | Configure un gestionnaire d’événements pour un événement spécifié. | **\<Event\>** et **\<ExtensionPoint\>** (de type Événements) |*Aucun.* |

##### <a name="ribbons-table"></a>Tableau « rubans »

Le tableau suivant mappe les propriétés enfants des objets enfants anonymes du tableau « rubans » aux éléments XML du manifeste actuel. 

|Propriété JSON|Objectif|Éléments XML|Commentaires|
|:-----|:-----|:-----|:-----|
| « contextes » | Spécifie les surfaces de commande que le complément personnalise. | divers éléments **\*CommandSurface**, tels que **PrimaryCommandSurface** et **MessageReadCommandSurface** |*Aucun.* |
| « onglets » | Configure les onglets du ruban personnalisé. | **\<CustomTab\>** | Les noms et la hiérarchie des propriétés descendantes de « onglets » correspondent étroitement aux descendants de **\<CustomTab\>**.  |

## <a name="sample-preview-json-manifest"></a>Exemple de manifeste JSON d’aperçu

Voici un exemple de manifeste JSON d’aperçu pour un complément.

```json
{
  "$schema": "https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/op/extensions/MicrosoftTeams.schema.json",
  "id": "00000000-0000-0000-0000-000000000000",
  "version": "1.0.0",
  "manifestVersion": "devPreview",
  "name": {
    "short": "Name of your app (<=30 chars)",
    "full": "Full name of app, if longer than 30 characters (<=100 chars)"
  },
  "description": {
    "short": "Short description of your app (<= 80 chars)",
    "full": "Full description of your app (<= 4000 chars)"
  },
  "icons": {
    "outline": "outline.png",
    "color": "color.png"
  },
  "accentColor": "#230201",
  "developer": {
    "name": "Contoso",
    "websiteUrl": "https://www.contoso.com",
    "privacyUrl": "https://www.contoso.com/privacy",
    "termsOfUseUrl": "https://www.contoso.com/servicesagreement"
  },
  "localizationInfo": {
    "defaultLanguageTag": "en-us",
    "additionalLanguages": [
      {
        "languageTag": "es-es",
        "file": "es-es.json"
      }
    ]
  },
  "webApplicationInfo": {
    "id": "00000000-0000-0000-0000-000000000000",
    "resource": "api://www.contoso.com/prodapp"
  },
  "authorization": {
    "permissions": {
      "resourceSpecific": [
        {
          "name": "Mailbox.ReadWrite.User",
          "type": "Delegated"
        }
      ]
    }
  },
  "extensions": [
    {
      "requirements": {
        "scopes": [ "mail" ],
        "capabilities": [
          {
            "name": "Mailbox", "minVersion": "1.1"
          }
        ]
      },
      "runtimes": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "id": "eventsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/events.html",
            "script": "https://contoso.com/events.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "onMessageSending",
              "type": "executeFunction"
            },
            {
              "id": "onNewMessageComposeCreated",
              "type": "executeFunction"
            }
          ]
        },
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.1"
              }
            ]
          },
          "id": "commandsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/commands.html",
            "script": "https://contoso.com/commands.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "action1",
              "type": "executeFunction"
            },
            {
              "id": "action2",
              "type": "executeFunction"
            },
            {
              "id": "action3",
              "type": "executeFunction"
            }
          ]
        }
      ],
      "ribbons": [
        {
          "contexts": [
            "mailCompose"
          ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    },
                    {
                      "id": "menu1",
                      "type": "menu",
                      "label": "My Menu",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "My Menu",
                        "description": "Menu with 2 actions"
                      },
                      "items": [
                        {
                          "id": "menuItem1",
                          "type": "menuItem",
                          "label": "Action 2",
                          "supertip": {
                            "title": "Action 2 Title",
                            "description": "Action 2 Description"
                          },
                          "actionId": "action2"
                        },
                        {
                          "id": "menuItem2",
                          "type": "menuItem",
                          "label": "Action 3",
                          "icons": [
                            {
                              "size": 16,
                              "file": "test_16.png"
                            },
                            {
                              "size": 32,
                              "file": "test_32.png"
                            },
                            {
                              "size": 80,
                              "file": "test_80.png"
                            }
                          ],
                          "supertip": {
                            "title": "Action 3 Title",
                            "description": "Action 3 Description"
                          },
                          "actionId": "action3"
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "contexts": [ "mailRead" ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ],
      "autoRunEvents": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "events": [
            {
              "type": "newMessageComposeCreated",
              "actionId": "onNewMessageComposeCreated"
            },
            {
              "type": "messageSending",
              "actionId": "onMessageSending",
              "options": {
                "sendMode": "promptUser"
              }
            }
          ]
        }
      ],
      "alternates": [
        {
          "requirements": {
            "scopes": [ "mail" ]
          },
          "prefer": {
            "comAddin": {
              "progId": "ContosoExtension"
            }
          },
          "hide": {
            "storeOfficeAddin": {
              "officeAddinId": "00000000-0000-0000-0000-000000000000",
              "assetId": "WA000000000"
            }
          }
        }
      ]
    }
  ]
}
```

## <a name="next-steps"></a>Prochaines étapes

- [Créez un complément Outlook avec un manifeste Teams (préversion).](../quickstarts/outlook-quickstart-json-manifest.md)