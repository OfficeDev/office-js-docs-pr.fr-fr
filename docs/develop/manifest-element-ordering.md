---
title: Comment trouver l’ordre approprié d’éléments manifeste
description: Découvrez comment trouver l’ordre correct dans lequel placer les éléments enfants dans un élément parent.
ms.date: 11/16/2018
localization_priority: Normal
ms.openlocfilehash: a7ec2e5b0dee5be651e4670effd86bc4acbac028
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268120"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>Comment trouver l’ordre approprié d’éléments manifeste

Les éléments XML dans le fichier manifeste d’un complément Office doivent être sous l’élément parent approprié *et* dans un ordre spécifique, par rapport à d’autres, sous le parent.

Le classement requis est spécifié dans les fichiers XSD dans le dossier [schémas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). Les fichiers XSD sont classés dans des sous-dossiers pour volet de tâches, contenu et compléments de courrier.

Par exemple, dans l’`<OfficeApp>`élément, le `<Id>`,`<Version>` ,`<ProviderName>` doit apparaître dans cet ordre. Si un élément `<AlternateId>` est ajouté, il doit être compris entre l’élément `<Id>` et `<Version>`. Votre manifeste ne sera pas valide et votre complément ne sera pas chargé, si un élément n’est pas dans l’ordre.

> [!NOTE]
> Le [validateur au sein de la boîte à outils Office](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-toolbox) utilise le même message d’erreur lorsqu’un élément est absent de l’ordre lorsqu’un élément est sous un parent incorrect. L’erreur indique que l’élément enfant n’est pas un enfant valide de l’élément parent. Si vous recevez un message d’erreur mais que la documentation de référence pour l’élément enfant indique qu’elle *est* valide pour le parent, alors le problème est probablement que l’enfant a été placé dans l’ordre incorrect.

Pour rechercher l’ordre correct pour les éléments enfants d’un élément parent donné, procédez comme suit. (C’est un processus simplifié, car les fichiers XSD sont relativement complexes. L’analyse entière des fichiers XSD est hors de l’étendue de ce document.)

1. Ouvrez le sous-dossier [schémas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) correspondant au type de complément que vous créez. 
2. Ouvrez le fichier XSD où l’élément parent est défini comme un type complexe. Si vous ne savez pas quel fichier a la définition, vous devrez peut-être répéter l’étape 3 sur plusieurs fichiers jusqu'à ce que vous trouviez.
3. Recherchez `<xs:complexType name="PARENT_ELEMENT">`, où PARENT_ELEMENT est le nom de l’élément parent.
4. À l’intérieur de la définition pour le PARENT_ELEMENT, il y a (généralement) un élément appelé `<xs:sequence>`. Voici la définition pour `<SuperTip>` de [TaskPaneAppVersionOverridesV1_0.xsd](https://raw.githubusercontent.com/OfficeDev/office-js-docs-pr/master/docs/overview/schemas/taskpane/TaskPaneAppVersionOverridesV1_0.xsd).

```xml
  <xs:complexType name="Supertip">
    <xs:annotation>
      <xs:documentation>
        Specifies the super tip for this control.
      </xs:documentation>
    </xs:annotation>
    <xs:sequence>
      <xs:element name="Title" type="bt:ShortResourceReference" minOccurs="1" maxOccurs="1" />
      <xs:element name="Description" type="bt:LongResourceReference" minOccurs="1" maxOccurs="1" />
    </xs:sequence>
  </xs:complexType>
```

Le `<xs:sequence>` répertorie les éléments enfants possibles *dans l’ordre dans lequel ils doivent apparaître*. Cette option ne signifie *pas* qu’ils sont tous sont obligatoires. Si la`minOccurs` valeur pour un élément enfant est **0**, alors l’élément enfant est facultatif. *Mais s’il apparaît, il doit être dans l’ordre spécifié par l’ `<xs:sequence>` élément*.

S’il n’y a aucun`<xs:sequence>` élément, ou qu’il *est* présent mais l’élément enfant n’est pas listé (même si la documentation de référence pour l’élément enfant indique qu’il *est* valide pour le parent) ; alors la définition de l’élément parent type complexe a été étendue avec des éléments enfants supplémentaires ailleurs dans le fichier XSD. Par exemple, la définition pour le `OfficeApp` type complexe ne répertorie pas `Requirements` comme enfant possible. Mais plus loin dans le fichier (au sein de la définition pour le `TaskPaneApp` type complexe), la définition de `OfficeApp` est prolongée et `Requirements` est ajoutée comme un enfant valide supplémentaire.

Pour trouver les définitions étendues procédez comme suit :

1. En commençant au haut du fichier, recherchez `<xs:extension base="PARENT_ELEMENT">`, où PARENT_ELEMENT est le nom de l’élément parent. Il existe peut-être plus d’une extension.
2. Rechercher l’extension pertinente pour le contexte dans lequel vous travaillez. Par exemple, le `OfficeApp` type complexe s’étend au sein des `ContentApp` et `MailApp`types complexes ainsi que dans le `TaskPaneApp` type complexe.

Chaque `<xs:extension base="PARENT_ELEMENT">` dans le fichier apparaît avec son propre `<xs:sequence>` qui contient des éléments enfants valides supplémentaires pour le parent. Les éléments enfants sur une liste étendue doivent toujours être *après* les éléments enfants dans la liste d’origine dans la définition de type complexe du parent.

## <a name="see-also"></a>Voir aussi

- [Référence de schéma pour les manifestes des compléments Office (version 1.1)](../develop/add-in-manifests.md)
