---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,1
description: Détails sur l’ensemble de conditions requises ExcelApi 1,1
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 921a67b4242150d767fdac057d21c6fc510d98b3
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772050"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Ensemble de conditions requises de l’API JavaScript pour Excel 1,1

L’API JavaScript 1.1 pour Excel est la première version de l’API. Il s’agit du seul ensemble de conditions requises spécifiques à Excel pris en charge par Excel 2016.

## <a name="api-list"></a>Liste des API

| Class | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[Calculate (calculationType: "recalculer \| " "Full \| " "FullRebuild")](/javascript/api/excel/excel.application#calculate-calculationtype-)|Recalcule tous les classeurs actuellement ouverts dans Excel.|
||[Calculate (calculationType: Excel. CalculationType)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Recalcule tous les classeurs actuellement ouverts dans Excel.|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|Renvoie le mode de calcul utilisé dans le classeur, tel que défini par les constantes dans Excel. CalculationMode. Les valeurs possibles sont `Automatic`les suivantes:, où Excel contrôle le recalcul; `AutomaticExceptTables`, où Excel contrôle le recalcul, mais ignore les modifications apportées aux tableaux; `Manual`, où le calcul est effectué lorsque l’utilisateur le demande.|
||[Set (propriétés: Excel. application)](/javascript/api/excel/excel.application#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ApplicationUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.application#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ApplicationData](/javascript/api/excel/excel.applicationdata)|[calculationMode](/javascript/api/excel/excel.applicationdata#calculationmode)|Renvoie le mode de calcul utilisé dans le classeur, tel que défini par les constantes dans Excel. CalculationMode. Les valeurs possibles sont `Automatic`les suivantes:, où Excel contrôle le recalcul; `AutomaticExceptTables`, où Excel contrôle le recalcul, mais ignore les modifications apportées aux tableaux; `Manual`, où le calcul est effectué lorsque l’utilisateur le demande.|
|[ApplicationLoadOptions](/javascript/api/excel/excel.applicationloadoptions)|[$all](/javascript/api/excel/excel.applicationloadoptions#$all)||
||[calculationMode](/javascript/api/excel/excel.applicationloadoptions#calculationmode)|Renvoie le mode de calcul utilisé dans le classeur, tel que défini par les constantes dans Excel. CalculationMode. Les valeurs possibles sont `Automatic`les suivantes:, où Excel contrôle le recalcul; `AutomaticExceptTables`, où Excel contrôle le recalcul, mais ignore les modifications apportées aux tableaux; `Manual`, où le calcul est effectué lorsque l’utilisateur le demande.|
|[ApplicationUpdateData](/javascript/api/excel/excel.applicationupdatedata)|[calculationMode](/javascript/api/excel/excel.applicationupdatedata#calculationmode)|Renvoie le mode de calcul utilisé dans le classeur, tel que défini par les constantes dans Excel. CalculationMode. Les valeurs possibles sont `Automatic`les suivantes:, où Excel contrôle le recalcul; `AutomaticExceptTables`, où Excel contrôle le recalcul, mais ignore les modifications apportées aux tableaux; `Manual`, où le calcul est effectué lorsque l’utilisateur le demande.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|Renvoie la plage représentée par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|Renvoie le tableau représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|Renvoie le texte représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
||[id](/javascript/api/excel/excel.binding#id)|Représente l’identificateur de liaison. En lecture seule.|
||[type](/javascript/api/excel/excel.binding#type)|Renvoie le type de la liaison. Pour plus d’informations, voir Excel. BindingType. En lecture seule.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|Obtient un objet de liaison par ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|Obtient un objet de liaison en fonction de sa position dans le tableau d’éléments.|
||[count](/javascript/api/excel/excel.bindingcollection#count)|Renvoie le nombre de liaisons de la collection. En lecture seule.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[BindingCollectionLoadOptions](/javascript/api/excel/excel.bindingcollectionloadoptions)|[$all](/javascript/api/excel/excel.bindingcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.bindingcollectionloadoptions#id)|Pour chaque élément de la collection: représente l’identificateur de liaison. En lecture seule.|
||[type](/javascript/api/excel/excel.bindingcollectionloadoptions#type)|Pour chaque élément de la collection: renvoie le type de la liaison. Pour plus d’informations, voir Excel. BindingType. En lecture seule.|
|[BindingData](/javascript/api/excel/excel.bindingdata)|[id](/javascript/api/excel/excel.bindingdata#id)|Représente l’identificateur de liaison. En lecture seule.|
||[type](/javascript/api/excel/excel.bindingdata#type)|Renvoie le type de la liaison. Pour plus d’informations, voir Excel. BindingType. En lecture seule.|
|[BindingLoadOptions](/javascript/api/excel/excel.bindingloadoptions)|[$all](/javascript/api/excel/excel.bindingloadoptions#$all)||
||[id](/javascript/api/excel/excel.bindingloadoptions#id)|Représente l’identificateur de liaison. En lecture seule.|
||[type](/javascript/api/excel/excel.bindingloadoptions#type)|Renvoie le type de la liaison. Pour plus d’informations, voir Excel. BindingType. En lecture seule.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|Supprime l’objet de graphique.|
||[height](/javascript/api/excel/excel.chart#height)|Représente la hauteur, exprimée en points, de l’objet de graphique.|
||[left](/javascript/api/excel/excel.chart#left)|Distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.chart#name)|Représente le nom d’un objet de graphique.|
||[ordonné](/javascript/api/excel/excel.chart#axes)|Représente les axes du graphique. En lecture seule.|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|Représente les étiquettes des données sur le graphique. En lecture seule.|
||[format](/javascript/api/excel/excel.chart#format)|Regroupe les propriétés de format de la zone de graphique. En lecture seule.|
||[Legend](/javascript/api/excel/excel.chart#legend)|Représente la légende du graphique. En lecture seule.|
||[série](/javascript/api/excel/excel.chart#series)|Représente une série ou une collection de séries dans le graphique. En lecture seule.|
||[title](/javascript/api/excel/excel.chart#title)|Représente le titre du graphique indiqué et comprend le texte, la visibilité, la position et la mise en forme du titre. En lecture seule.|
||[Set (propriétés: Excel. Chart)](/javascript/api/excel/excel.chart#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chart#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[setData (sourceData: Range, seriesBy?: "auto" \| "Columns \| " "rows")](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Redéfinit les données sources du graphique.|
||[setData (sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Redéfinit les données sources du graphique.|
||[setPosition (startCell: chaîne \| de plage, endCell?: \| chaîne de plage)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|Positionne le graphique par rapport aux cellules dans la feuille de calcul.|
||[top](/javascript/api/excel/excel.chart#top)|Représente la distance, en points, entre le bord supérieur de l’objet et la partie supérieure de la ligne 1 (sur une feuille de calcul) ou le haut de la zone de graphique (sur un graphique).|
||[width](/javascript/api/excel/excel.chart#width)|Représente la largeur, en points, de l’objet de graphique.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan. En lecture seule.|
||[police](/javascript/api/excel/excel.chartareaformat#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) de l’objet. En lecture seule.|
||[Set (propriétés: Excel. ChartAreaFormat)](/javascript/api/excel/excel.chartareaformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartAreaFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartareaformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[police](/javascript/api/excel/excel.chartareaformatdata#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) de l’objet. En lecture seule.|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartareaformatloadoptions#$all)||
||[police](/javascript/api/excel/excel.chartareaformatloadoptions#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) de l’objet.|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[police](/javascript/api/excel/excel.chartareaformatupdatedata#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) de l’objet.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|Représente l’axe des abscisses d’un graphique. En lecture seule.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesaxis)|Représente l’axe de séries d’un graphique 3D. En lecture seule.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueaxis)|Représente l’axe des ordonnées. En lecture seule.|
||[Set (propriétés: Excel. ChartAxes)](/javascript/api/excel/excel.chartaxes#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartAxesUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartaxes#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartAxesData](/javascript/api/excel/excel.chartaxesdata)|[categoryAxis](/javascript/api/excel/excel.chartaxesdata#categoryaxis)|Représente l’axe des abscisses d’un graphique. En lecture seule.|
||[seriesAxis](/javascript/api/excel/excel.chartaxesdata#seriesaxis)|Représente l’axe de séries d’un graphique 3D. En lecture seule.|
||[valueAxis](/javascript/api/excel/excel.chartaxesdata#valueaxis)|Représente l’axe des ordonnées. En lecture seule.|
|[ChartAxesLoadOptions](/javascript/api/excel/excel.chartaxesloadoptions)|[$all](/javascript/api/excel/excel.chartaxesloadoptions#$all)||
||[categoryAxis](/javascript/api/excel/excel.chartaxesloadoptions#categoryaxis)|Représente l’axe des abscisses d’un graphique.|
||[seriesAxis](/javascript/api/excel/excel.chartaxesloadoptions#seriesaxis)|Représente l’axe de séries d’un graphique 3D.|
||[valueAxis](/javascript/api/excel/excel.chartaxesloadoptions#valueaxis)|Représente l’axe des ordonnées.|
|[ChartAxesUpdateData](/javascript/api/excel/excel.chartaxesupdatedata)|[categoryAxis](/javascript/api/excel/excel.chartaxesupdatedata#categoryaxis)|Représente l’axe des abscisses d’un graphique.|
||[seriesAxis](/javascript/api/excel/excel.chartaxesupdatedata#seriesaxis)|Représente l’axe de séries d’un graphique 3D.|
||[valueAxis](/javascript/api/excel/excel.chartaxesupdatedata#valueaxis)|Représente l’axe des ordonnées.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|Représente l’intervalle entre deux graduations principales. Peut être défini sur une valeur numérique ou une chaîne vide.  La valeur renvoyée est toujours un nombre.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Représente la valeur maximale sur l’axe des ordonnées.  Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique).  La valeur renvoyée est toujours un nombre.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Représente la valeur minimale sur l’axe des ordonnées. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|Représente l’intervalle entre deux graduations secondaires. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Représente la mise en forme d’un objet de graphique, à savoir le format des lignes et de la police. En lecture seule.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage principal de l’axe spécifié. En lecture seule.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage secondaire de l’axe spécifié. En lecture seule.|
||[title](/javascript/api/excel/excel.chartaxis#title)|Représente le titre de l’axe. En lecture seule.|
||[Set (propriétés: Excel. ChartAxis)](/javascript/api/excel/excel.chartaxis#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartAxisUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartaxis#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[format](/javascript/api/excel/excel.chartaxisdata#format)|Représente la mise en forme d’un objet de graphique, à savoir le format des lignes et de la police. En lecture seule.|
||[majorGridlines](/javascript/api/excel/excel.chartaxisdata#majorgridlines)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage principal de l’axe spécifié. En lecture seule.|
||[majorUnit](/javascript/api/excel/excel.chartaxisdata#majorunit)|Représente l’intervalle entre deux graduations principales. Peut être défini sur une valeur numérique ou une chaîne vide.  La valeur renvoyée est toujours un nombre.|
||[maximum](/javascript/api/excel/excel.chartaxisdata#maximum)|Représente la valeur maximale sur l’axe des ordonnées.  Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique).  La valeur renvoyée est toujours un nombre.|
||[minimum](/javascript/api/excel/excel.chartaxisdata#minimum)|Représente la valeur minimale sur l’axe des ordonnées. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
||[minorGridlines](/javascript/api/excel/excel.chartaxisdata#minorgridlines)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage secondaire de l’axe spécifié. En lecture seule.|
||[minorUnit](/javascript/api/excel/excel.chartaxisdata#minorunit)|Représente l’intervalle entre deux graduations secondaires. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
||[title](/javascript/api/excel/excel.chartaxisdata#title)|Représente le titre de l’axe. En lecture seule.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[police](/javascript/api/excel/excel.chartaxisformat#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’un élément d’axe de graphique. En lecture seule.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Représente le format des lignes du graphique. En lecture seule.|
||[Set (propriétés: Excel. ChartAxisFormat)](/javascript/api/excel/excel.chartaxisformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartAxisFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartaxisformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartAxisFormatData](/javascript/api/excel/excel.chartaxisformatdata)|[police](/javascript/api/excel/excel.chartaxisformatdata#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’un élément d’axe de graphique. En lecture seule.|
||[line](/javascript/api/excel/excel.chartaxisformatdata#line)|Représente le format des lignes du graphique. En lecture seule.|
|[ChartAxisFormatLoadOptions](/javascript/api/excel/excel.chartaxisformatloadoptions)|[$all](/javascript/api/excel/excel.chartaxisformatloadoptions#$all)||
||[police](/javascript/api/excel/excel.chartaxisformatloadoptions#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’un élément d’axe de graphique.|
||[line](/javascript/api/excel/excel.chartaxisformatloadoptions#line)|Représente le format des lignes du graphique.|
|[ChartAxisFormatUpdateData](/javascript/api/excel/excel.chartaxisformatupdatedata)|[police](/javascript/api/excel/excel.chartaxisformatupdatedata#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’un élément d’axe de graphique.|
||[line](/javascript/api/excel/excel.chartaxisformatupdatedata#line)|Représente le format des lignes du graphique.|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[$all](/javascript/api/excel/excel.chartaxisloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartaxisloadoptions#format)|Représente la mise en forme d’un objet de graphique, à savoir le format des lignes et de la police.|
||[majorGridlines](/javascript/api/excel/excel.chartaxisloadoptions#majorgridlines)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage principal de l’axe spécifié.|
||[majorUnit](/javascript/api/excel/excel.chartaxisloadoptions#majorunit)|Représente l’intervalle entre deux graduations principales. Peut être défini sur une valeur numérique ou une chaîne vide.  La valeur renvoyée est toujours un nombre.|
||[maximum](/javascript/api/excel/excel.chartaxisloadoptions#maximum)|Représente la valeur maximale sur l’axe des ordonnées.  Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique).  La valeur renvoyée est toujours un nombre.|
||[minimum](/javascript/api/excel/excel.chartaxisloadoptions#minimum)|Représente la valeur minimale sur l’axe des ordonnées. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
||[minorGridlines](/javascript/api/excel/excel.chartaxisloadoptions#minorgridlines)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage secondaire de l’axe spécifié.|
||[minorUnit](/javascript/api/excel/excel.chartaxisloadoptions#minorunit)|Représente l’intervalle entre deux graduations secondaires. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
||[title](/javascript/api/excel/excel.chartaxisloadoptions#title)|Représente le titre de l’axe.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Représente le format du titre d’un axe de graphique. En lecture seule.|
||[Set (propriétés: Excel. ChartAxisTitle)](/javascript/api/excel/excel.chartaxistitle#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartAxisTitleUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartaxistitle#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Représente le titre de l’axe.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Valeur booléenne qui spécifie la visibilité d’un titre d’axe.|
|[ChartAxisTitleData](/javascript/api/excel/excel.chartaxistitledata)|[format](/javascript/api/excel/excel.chartaxistitledata#format)|Représente le format du titre d’un axe de graphique. En lecture seule.|
||[text](/javascript/api/excel/excel.chartaxistitledata#text)|Représente le titre de l’axe.|
||[visible](/javascript/api/excel/excel.chartaxistitledata#visible)|Valeur booléenne qui spécifie la visibilité d’un titre d’axe.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[police](/javascript/api/excel/excel.chartaxistitleformat#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc., de l’objet de titre d’axe de graphique. En lecture seule.|
||[Set (propriétés: Excel. ChartAxisTitleFormat)](/javascript/api/excel/excel.chartaxistitleformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartAxisTitleFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartaxistitleformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartAxisTitleFormatData](/javascript/api/excel/excel.chartaxistitleformatdata)|[police](/javascript/api/excel/excel.chartaxistitleformatdata#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc., de l’objet de titre d’axe de graphique. En lecture seule.|
|[ChartAxisTitleFormatLoadOptions](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[$all](/javascript/api/excel/excel.chartaxistitleformatloadoptions#$all)||
||[police](/javascript/api/excel/excel.chartaxistitleformatloadoptions#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc., de l’objet de titre d’axe de graphique.|
|[ChartAxisTitleFormatUpdateData](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[police](/javascript/api/excel/excel.chartaxistitleformatupdatedata#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc., de l’objet de titre d’axe de graphique.|
|[ChartAxisTitleLoadOptions](/javascript/api/excel/excel.chartaxistitleloadoptions)|[$all](/javascript/api/excel/excel.chartaxistitleloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartaxistitleloadoptions#format)|Représente le format du titre d’un axe de graphique.|
||[text](/javascript/api/excel/excel.chartaxistitleloadoptions#text)|Représente le titre de l’axe.|
||[visible](/javascript/api/excel/excel.chartaxistitleloadoptions#visible)|Valeur booléenne qui spécifie la visibilité d’un titre d’axe.|
|[ChartAxisTitleUpdateData](/javascript/api/excel/excel.chartaxistitleupdatedata)|[format](/javascript/api/excel/excel.chartaxistitleupdatedata#format)|Représente le format du titre d’un axe de graphique.|
||[text](/javascript/api/excel/excel.chartaxistitleupdatedata#text)|Représente le titre de l’axe.|
||[visible](/javascript/api/excel/excel.chartaxistitleupdatedata#visible)|Valeur booléenne qui spécifie la visibilité d’un titre d’axe.|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[format](/javascript/api/excel/excel.chartaxisupdatedata#format)|Représente la mise en forme d’un objet de graphique, à savoir le format des lignes et de la police.|
||[majorGridlines](/javascript/api/excel/excel.chartaxisupdatedata#majorgridlines)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage principal de l’axe spécifié.|
||[majorUnit](/javascript/api/excel/excel.chartaxisupdatedata#majorunit)|Représente l’intervalle entre deux graduations principales. Peut être défini sur une valeur numérique ou une chaîne vide.  La valeur renvoyée est toujours un nombre.|
||[maximum](/javascript/api/excel/excel.chartaxisupdatedata#maximum)|Représente la valeur maximale sur l’axe des ordonnées.  Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique).  La valeur renvoyée est toujours un nombre.|
||[minimum](/javascript/api/excel/excel.chartaxisupdatedata#minimum)|Représente la valeur minimale sur l’axe des ordonnées. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
||[minorGridlines](/javascript/api/excel/excel.chartaxisupdatedata#minorgridlines)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage secondaire de l’axe spécifié.|
||[minorUnit](/javascript/api/excel/excel.chartaxisupdatedata#minorunit)|Représente l’intervalle entre deux graduations secondaires. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
||[title](/javascript/api/excel/excel.chartaxisupdatedata#title)|Représente le titre de l’axe.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[Add (type: "non valide \| " "ColumnClustered \| " "ColumnStacked \| " "ColumnStacked100 \| " "3DColumnClustered \| " "3DColumnStacked \| " "3DColumnStacked100 \| " "BarClustered \| " "BarStacked" \| "BarStacked100" \| "3DBarClustered" \| "3DBarStacked" \| "3DBarStacked100" \| "LineStacked" \| "LineStacked100" \| \| "LineMarkers" "LineMarkersStacked" \| " LineMarkersStacked100 " \| " PieOfPie " \| " PieExploded " \| " 3DPieExploded " \| " BarOfPie " \| " XYScatterSmooth " \| " XYScatterSmoothNoMarkers " \| " XYScatterLines " \| " XYScatterLinesNoMarkers " \| " AreaStacked " \| " AreaStacked100 " \| " 3DAreaStacked " \| " 3DAreaStacked100 " \| " DoughnutExploded " \| " RadarMarkers " \| " RadarFilled " \| " Surface " \| " SurfaceWireframe " \| " SurfaceTopView " \| " SurfaceTopViewWireframe " \| " Bubble " \| " Bubble3DEffect " \| " StockHLC " \| " StockOHLC " \| " StockVHLC " \| " StockVOHLC " \| " CylinderColClustered " \| " CylinderColStacked " \| " CylinderColStacked100 " \| " CylinderBarClustered " \| " CylinderBarStacked " \| " CylinderBarStacked100 " \| " CylinderCol " \| " ConeColClustered " \| " ConeColStacked " \| " ConeColStacked100 " \| " ConeBarClustered " \| " ConeBarStacked " \| " ConeBarStacked100 " \| " ConeCol " \| " PyramidColClustered " \| " PyramidColStacked " \| " PyramidColStacked100 " \| " PyramidBarClustered " \| " PyramidBarStacked " \| " PyramidBarStacked100 " \| " PyramidCol " \| " 3DColumn " \| "Line" \| "3DLine" \| "3DPie" \| "Pie" \| "XYScatter" \| "3DArea" \| "area" \| "Pie" \| "radar" \| "histogramme" \| "Boxwhisker" \| " Pareto \| "" RegionMap \| "" TreeMap \| "" chutes \| "" cascade " \| " soleil "" entonnoir ", DonnéesSources: Range, seriesBy?: \| " auto " \| " Columns "" rows ")](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Crée un graphique.|
||[Add (type: Excel. ChartType, sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Crée un graphique.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|Extrait un graphique en fonction de sa position dans la collection.|
||[count](/javascript/api/excel/excel.chartcollection#count)|Renvoie le nombre de graphiques dans la feuille de calcul. En lecture seule.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[$all](/javascript/api/excel/excel.chartcollectionloadoptions#$all)||
||[ordonné](/javascript/api/excel/excel.chartcollectionloadoptions#axes)|Pour chaque élément de la collection: représente les axes de graphique.|
||[dataLabels](/javascript/api/excel/excel.chartcollectionloadoptions#datalabels)|Pour chaque élément de la collection: représente l’élément DataLabels sur le graphique.|
||[format](/javascript/api/excel/excel.chartcollectionloadoptions#format)|Pour chaque élément de la collection: encapsule les propriétés de format de la zone de graphique.|
||[height](/javascript/api/excel/excel.chartcollectionloadoptions#height)|Pour chaque élément de la collection: représente la hauteur, exprimée en points, de l’objet Chart.|
||[left](/javascript/api/excel/excel.chartcollectionloadoptions#left)|Pour chaque élément de la collection: distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[Legend](/javascript/api/excel/excel.chartcollectionloadoptions#legend)|Pour chaque élément de la collection: représente la légende du graphique.|
||[name](/javascript/api/excel/excel.chartcollectionloadoptions#name)|Pour chaque élément de la collection: représente le nom d’un objet de graphique.|
||[séquence](/javascript/api/excel/excel.chartcollectionloadoptions#series)|Pour chaque élément de la collection: représente une seule série ou collection de séries dans le graphique.|
||[title](/javascript/api/excel/excel.chartcollectionloadoptions#title)|Pour chaque élément de la collection: représente le titre du graphique spécifié, y compris le texte, la visibilité, la position et la mise en forme du titre.|
||[top](/javascript/api/excel/excel.chartcollectionloadoptions#top)|Pour chaque élément de la collection: représente la distance en points entre le bord supérieur de l’objet et le haut de ligne 1 (dans une feuille de calcul) ou le haut de la zone de graphique (dans un graphique).|
||[width](/javascript/api/excel/excel.chartcollectionloadoptions#width)|Pour chaque élément de la collection: représente la largeur, exprimée en points, de l’objet de graphique.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[ordonné](/javascript/api/excel/excel.chartdata#axes)|Représente les axes du graphique. En lecture seule.|
||[dataLabels](/javascript/api/excel/excel.chartdata#datalabels)|Représente les étiquettes des données sur le graphique. En lecture seule.|
||[format](/javascript/api/excel/excel.chartdata#format)|Regroupe les propriétés de format de la zone de graphique. En lecture seule.|
||[height](/javascript/api/excel/excel.chartdata#height)|Représente la hauteur, exprimée en points, de l’objet de graphique.|
||[left](/javascript/api/excel/excel.chartdata#left)|Distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[Legend](/javascript/api/excel/excel.chartdata#legend)|Représente la légende du graphique. En lecture seule.|
||[name](/javascript/api/excel/excel.chartdata#name)|Représente le nom d’un objet de graphique.|
||[séquence](/javascript/api/excel/excel.chartdata#series)|Représente une série ou une collection de séries dans le graphique. En lecture seule.|
||[title](/javascript/api/excel/excel.chartdata#title)|Représente le titre du graphique indiqué et comprend le texte, la visibilité, la position et la mise en forme du titre. En lecture seule.|
||[top](/javascript/api/excel/excel.chartdata#top)|Représente la distance, en points, entre le bord supérieur de l’objet et la partie supérieure de la ligne 1 (sur une feuille de calcul) ou le haut de la zone de graphique (sur un graphique).|
||[width](/javascript/api/excel/excel.chartdata#width)|Représente la largeur, en points, de l’objet de graphique.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Représente le format de remplissage de l’étiquette de données. En lecture seule.|
||[police](/javascript/api/excel/excel.chartdatalabelformat#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’une étiquette de données de graphique. En lecture seule.|
||[Set (propriétés: Excel. ChartDataLabelFormat)](/javascript/api/excel/excel.chartdatalabelformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartDataLabelFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartdatalabelformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartDataLabelFormatData](/javascript/api/excel/excel.chartdatalabelformatdata)|[police](/javascript/api/excel/excel.chartdatalabelformatdata#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’une étiquette de données de graphique. En lecture seule.|
|[ChartDataLabelFormatLoadOptions](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelformatloadoptions#$all)||
||[police](/javascript/api/excel/excel.chartdatalabelformatloadoptions#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’une étiquette de données de graphique.|
|[ChartDataLabelFormatUpdateData](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[police](/javascript/api/excel/excel.chartdatalabelformatupdatedata#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’une étiquette de données de graphique.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Pour plus d’informations, voir Excel. ChartDataLabelPosition.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Représente le format des étiquettes de données du graphique, à savoir le format de remplissage et de la police. En lecture seule.|
||[para](/javascript/api/excel/excel.chartdatalabels#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[Set (propriétés: Excel. ChartDataLabels)](/javascript/api/excel/excel.chartdatalabels#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartDataLabelsUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartdatalabels#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[format](/javascript/api/excel/excel.chartdatalabelsdata#format)|Représente le format des étiquettes de données du graphique, à savoir le format de remplissage et de la police. En lecture seule.|
||[position](/javascript/api/excel/excel.chartdatalabelsdata#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Pour plus d’informations, voir Excel. ChartDataLabelPosition.|
||[para](/javascript/api/excel/excel.chartdatalabelsdata#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsdata#showbubblesize)|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsdata#showcategoryname)|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabelsdata#showlegendkey)|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsdata#showpercentage)|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsdata#showseriesname)|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|
||[showValue](/javascript/api/excel/excel.chartdatalabelsdata#showvalue)|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelsloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartdatalabelsloadoptions#format)|Représente le format des étiquettes de données du graphique, à savoir le format de remplissage et de la police.|
||[position](/javascript/api/excel/excel.chartdatalabelsloadoptions#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Pour plus d’informations, voir Excel. ChartDataLabelPosition.|
||[para](/javascript/api/excel/excel.chartdatalabelsloadoptions#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsloadoptions#showbubblesize)|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsloadoptions#showcategoryname)|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabelsloadoptions#showlegendkey)|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsloadoptions#showpercentage)|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsloadoptions#showseriesname)|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|
||[showValue](/javascript/api/excel/excel.chartdatalabelsloadoptions#showvalue)|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[format](/javascript/api/excel/excel.chartdatalabelsupdatedata#format)|Représente le format des étiquettes de données du graphique, à savoir le format de remplissage et de la police.|
||[position](/javascript/api/excel/excel.chartdatalabelsupdatedata#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Pour plus d’informations, voir Excel. ChartDataLabelPosition.|
||[para](/javascript/api/excel/excel.chartdatalabelsupdatedata#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsupdatedata#showbubblesize)|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsupdatedata#showcategoryname)|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabelsupdatedata#showlegendkey)|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsupdatedata#showpercentage)|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsupdatedata#showseriesname)|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|
||[showValue](/javascript/api/excel/excel.chartdatalabelsupdatedata#showvalue)|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|Supprime la couleur de remplissage d’un élément de graphique.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|Définit le format de remplissage d’un élément de graphique sur une couleur unie.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.chartfont#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.chartfont#name)|Nom de la police (par exemple « Calibri »)|
||[Set (propriétés: Excel. ChartFont)](/javascript/api/excel/excel.chartfont#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartFontUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartfont#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[size](/javascript/api/excel/excel.chartfont#size)|Taille de la police (par exemple, 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. ChartUnderlineStyle.|
|[ChartFontData](/javascript/api/excel/excel.chartfontdata)|[bold](/javascript/api/excel/excel.chartfontdata#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.chartfontdata#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.chartfontdata#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.chartfontdata#name)|Nom de la police (par exemple « Calibri »)|
||[size](/javascript/api/excel/excel.chartfontdata#size)|Taille de la police (par exemple, 11)|
||[underline](/javascript/api/excel/excel.chartfontdata#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. ChartUnderlineStyle.|
|[ChartFontLoadOptions](/javascript/api/excel/excel.chartfontloadoptions)|[$all](/javascript/api/excel/excel.chartfontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.chartfontloadoptions#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.chartfontloadoptions#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.chartfontloadoptions#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.chartfontloadoptions#name)|Nom de la police (par exemple « Calibri »)|
||[size](/javascript/api/excel/excel.chartfontloadoptions#size)|Taille de la police (par exemple, 11)|
||[underline](/javascript/api/excel/excel.chartfontloadoptions#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. ChartUnderlineStyle.|
|[ChartFontUpdateData](/javascript/api/excel/excel.chartfontupdatedata)|[bold](/javascript/api/excel/excel.chartfontupdatedata#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.chartfontupdatedata#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.chartfontupdatedata#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.chartfontupdatedata#name)|Nom de la police (par exemple « Calibri »)|
||[size](/javascript/api/excel/excel.chartfontupdatedata#size)|Taille de la police (par exemple, 11)|
||[underline](/javascript/api/excel/excel.chartfontupdatedata#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. ChartUnderlineStyle.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Représente le format du quadrillage de graphique. En lecture seule.|
||[Set (propriétés: Excel. ChartGridlines)](/javascript/api/excel/excel.chartgridlines#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartGridlinesUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartgridlines#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Valeur booléenne indiquant si les lignes de quadrillage de l’axe sont visibles ou non.|
|[ChartGridlinesData](/javascript/api/excel/excel.chartgridlinesdata)|[format](/javascript/api/excel/excel.chartgridlinesdata#format)|Représente le format du quadrillage de graphique. En lecture seule.|
||[visible](/javascript/api/excel/excel.chartgridlinesdata#visible)|Valeur booléenne indiquant si les lignes de quadrillage de l’axe sont visibles ou non.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Représente le format des lignes du graphique. En lecture seule.|
||[Set (propriétés: Excel. ChartGridlinesFormat)](/javascript/api/excel/excel.chartgridlinesformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartGridlinesFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartgridlinesformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartGridlinesFormatData](/javascript/api/excel/excel.chartgridlinesformatdata)|[line](/javascript/api/excel/excel.chartgridlinesformatdata#line)|Représente le format des lignes du graphique. En lecture seule.|
|[ChartGridlinesFormatLoadOptions](/javascript/api/excel/excel.chartgridlinesformatloadoptions)|[$all](/javascript/api/excel/excel.chartgridlinesformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.chartgridlinesformatloadoptions#line)|Représente le format des lignes du graphique.|
|[ChartGridlinesFormatUpdateData](/javascript/api/excel/excel.chartgridlinesformatupdatedata)|[line](/javascript/api/excel/excel.chartgridlinesformatupdatedata#line)|Représente le format des lignes du graphique.|
|[ChartGridlinesLoadOptions](/javascript/api/excel/excel.chartgridlinesloadoptions)|[$all](/javascript/api/excel/excel.chartgridlinesloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartgridlinesloadoptions#format)|Représente le format du quadrillage de graphique.|
||[visible](/javascript/api/excel/excel.chartgridlinesloadoptions#visible)|Valeur booléenne indiquant si les lignes de quadrillage de l’axe sont visibles ou non.|
|[ChartGridlinesUpdateData](/javascript/api/excel/excel.chartgridlinesupdatedata)|[format](/javascript/api/excel/excel.chartgridlinesupdatedata#format)|Représente le format du quadrillage de graphique.|
||[visible](/javascript/api/excel/excel.chartgridlinesupdatedata#visible)|Valeur booléenne indiquant si les lignes de quadrillage de l’axe sont visibles ou non.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|Valeur booléenne indiquant si la légende du graphique doit chevaucher le corps principal du graphique.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Représente la position de la légende sur le graphique. Pour plus d’informations, voir Excel. ChartLegendPosition.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Représente le format d’une légende de graphique, à savoir le format du remplissage et de la police. En lecture seule.|
||[Set (propriétés: Excel. ChartLegend)](/javascript/api/excel/excel.chartlegend#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartLegendUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartlegend#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Valeur booléenne qui représente la visibilité d’un objet ChartLegend.|
|[ChartLegendData](/javascript/api/excel/excel.chartlegenddata)|[format](/javascript/api/excel/excel.chartlegenddata#format)|Représente le format d’une légende de graphique, à savoir le format du remplissage et de la police. En lecture seule.|
||[overlay](/javascript/api/excel/excel.chartlegenddata#overlay)|Valeur booléenne indiquant si la légende du graphique doit chevaucher le corps principal du graphique.|
||[position](/javascript/api/excel/excel.chartlegenddata#position)|Représente la position de la légende sur le graphique. Pour plus d’informations, voir Excel. ChartLegendPosition.|
||[visible](/javascript/api/excel/excel.chartlegenddata#visible)|Valeur booléenne qui représente la visibilité d’un objet ChartLegend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan. En lecture seule.|
||[police](/javascript/api/excel/excel.chartlegendformat#font)|Représente les attributs de police, tels que le nom de police, la taille de police, la couleur, etc., d’une légende de graphique. En lecture seule.|
||[Set (propriétés: Excel. ChartLegendFormat)](/javascript/api/excel/excel.chartlegendformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartLegendFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartlegendformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartLegendFormatData](/javascript/api/excel/excel.chartlegendformatdata)|[police](/javascript/api/excel/excel.chartlegendformatdata#font)|Représente les attributs de police, tels que le nom de police, la taille de police, la couleur, etc., d’une légende de graphique. En lecture seule.|
|[ChartLegendFormatLoadOptions](/javascript/api/excel/excel.chartlegendformatloadoptions)|[$all](/javascript/api/excel/excel.chartlegendformatloadoptions#$all)||
||[police](/javascript/api/excel/excel.chartlegendformatloadoptions#font)|Représente les attributs de police, tels que le nom de police, la taille de police, la couleur, etc., d’une légende de graphique.|
|[ChartLegendFormatUpdateData](/javascript/api/excel/excel.chartlegendformatupdatedata)|[police](/javascript/api/excel/excel.chartlegendformatupdatedata#font)|Représente les attributs de police, tels que le nom de police, la taille de police, la couleur, etc., d’une légende de graphique.|
|[ChartLegendLoadOptions](/javascript/api/excel/excel.chartlegendloadoptions)|[$all](/javascript/api/excel/excel.chartlegendloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartlegendloadoptions#format)|Représente le format d’une légende de graphique, à savoir le format du remplissage et de la police.|
||[overlay](/javascript/api/excel/excel.chartlegendloadoptions#overlay)|Valeur booléenne indiquant si la légende du graphique doit chevaucher le corps principal du graphique.|
||[position](/javascript/api/excel/excel.chartlegendloadoptions#position)|Représente la position de la légende sur le graphique. Pour plus d’informations, voir Excel. ChartLegendPosition.|
||[visible](/javascript/api/excel/excel.chartlegendloadoptions#visible)|Valeur booléenne qui représente la visibilité d’un objet ChartLegend.|
|[ChartLegendUpdateData](/javascript/api/excel/excel.chartlegendupdatedata)|[format](/javascript/api/excel/excel.chartlegendupdatedata#format)|Représente le format d’une légende de graphique, à savoir le format du remplissage et de la police.|
||[overlay](/javascript/api/excel/excel.chartlegendupdatedata#overlay)|Valeur booléenne indiquant si la légende du graphique doit chevaucher le corps principal du graphique.|
||[position](/javascript/api/excel/excel.chartlegendupdatedata#position)|Représente la position de la légende sur le graphique. Pour plus d’informations, voir Excel. ChartLegendPosition.|
||[visible](/javascript/api/excel/excel.chartlegendupdatedata#visible)|Valeur booléenne qui représente la visibilité d’un objet ChartLegend.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|Désactiver le format de ligne d’un élément de graphique.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|Code couleur HTML qui représente la couleur des lignes dans le graphique.|
||[Set (propriétés: Excel. ChartLineFormat)](/javascript/api/excel/excel.chartlineformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartLineFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartlineformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartLineFormatData](/javascript/api/excel/excel.chartlineformatdata)|[color](/javascript/api/excel/excel.chartlineformatdata#color)|Code couleur HTML qui représente la couleur des lignes dans le graphique.|
|[ChartLineFormatLoadOptions](/javascript/api/excel/excel.chartlineformatloadoptions)|[$all](/javascript/api/excel/excel.chartlineformatloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartlineformatloadoptions#color)|Code couleur HTML qui représente la couleur des lignes dans le graphique.|
|[ChartLineFormatUpdateData](/javascript/api/excel/excel.chartlineformatupdatedata)|[color](/javascript/api/excel/excel.chartlineformatupdatedata#color)|Code couleur HTML qui représente la couleur des lignes dans le graphique.|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[$all](/javascript/api/excel/excel.chartloadoptions#$all)||
||[ordonné](/javascript/api/excel/excel.chartloadoptions#axes)|Représente les axes du graphique.|
||[dataLabels](/javascript/api/excel/excel.chartloadoptions#datalabels)|Représente les étiquettes des données sur le graphique.|
||[format](/javascript/api/excel/excel.chartloadoptions#format)|Regroupe les propriétés de format de la zone de graphique.|
||[height](/javascript/api/excel/excel.chartloadoptions#height)|Représente la hauteur, exprimée en points, de l’objet de graphique.|
||[left](/javascript/api/excel/excel.chartloadoptions#left)|Distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[Legend](/javascript/api/excel/excel.chartloadoptions#legend)|Représente la légende du graphique.|
||[name](/javascript/api/excel/excel.chartloadoptions#name)|Représente le nom d’un objet de graphique.|
||[séquence](/javascript/api/excel/excel.chartloadoptions#series)|Représente une série ou une collection de séries dans le graphique.|
||[title](/javascript/api/excel/excel.chartloadoptions#title)|Représente le titre du graphique indiqué et comprend le texte, la visibilité, la position et la mise en forme du titre.|
||[top](/javascript/api/excel/excel.chartloadoptions#top)|Représente la distance, en points, entre le bord supérieur de l’objet et la partie supérieure de la ligne 1 (sur une feuille de calcul) ou le haut de la zone de graphique (sur un graphique).|
||[width](/javascript/api/excel/excel.chartloadoptions#width)|Représente la largeur, en points, de l’objet de graphique.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Regroupe les propriétés de format d’un point d’un graphique. En lecture seule.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Renvoie la valeur d’un point du graphique. En lecture seule.|
||[Set (propriétés: Excel. ChartPoint)](/javascript/api/excel/excel.chartpoint#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartPointUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartpoint#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartPointData](/javascript/api/excel/excel.chartpointdata)|[format](/javascript/api/excel/excel.chartpointdata#format)|Regroupe les propriétés de format d’un point d’un graphique. En lecture seule.|
||[value](/javascript/api/excel/excel.chartpointdata#value)|Renvoie la valeur d’un point du graphique. En lecture seule.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Représente le format de remplissage d’un graphique, qui inclut des informations de mise en forme de l’arrière-plan. En lecture seule.|
||[Set (propriétés: Excel. ChartPointFormat)](/javascript/api/excel/excel.chartpointformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartPointFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartpointformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartPointFormatLoadOptions](/javascript/api/excel/excel.chartpointformatloadoptions)|[$all](/javascript/api/excel/excel.chartpointformatloadoptions#$all)||
|[ChartPointLoadOptions](/javascript/api/excel/excel.chartpointloadoptions)|[$all](/javascript/api/excel/excel.chartpointloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartpointloadoptions#format)|Regroupe les propriétés de format d’un point d’un graphique.|
||[value](/javascript/api/excel/excel.chartpointloadoptions#value)|Renvoie la valeur d’un point du graphique. En lecture seule.|
|[ChartPointUpdateData](/javascript/api/excel/excel.chartpointupdatedata)|[format](/javascript/api/excel/excel.chartpointupdatedata#format)|Regroupe les propriétés de format d’un point d’un graphique.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|Extrait un point en fonction de sa position dans la série.|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|Renvoie le nombre de points de graphique dans la série. En lecture seule.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartPointsCollectionLoadOptions](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[$all](/javascript/api/excel/excel.chartpointscollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartpointscollectionloadoptions#format)|Pour chaque élément de la collection: encapsule le point de graphique propriétés de format.|
||[value](/javascript/api/excel/excel.chartpointscollectionloadoptions#value)|Pour chaque élément de la collection: renvoie la valeur d’un point de graphique. En lecture seule.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Représente le nom d’une série dans un graphique.|
||[format](/javascript/api/excel/excel.chartseries#format)|Représente le format d’une série de graphique, à savoir le format de remplissage et des lignes. En lecture seule.|
||[pointe](/javascript/api/excel/excel.chartseries#points)|Représente la collection de tous les points de la série. En lecture seule.|
||[Set (propriétés: Excel. ChartSeries)](/javascript/api/excel/excel.chartseries#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartSeriesUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartseries#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|Extrait une série en fonction de sa position dans la collection.|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|Renvoie le nombre de séries de la collection. En lecture seule.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[$all](/javascript/api/excel/excel.chartseriescollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartseriescollectionloadoptions#format)|Pour chaque élément de la collection: représente la mise en forme d’une série de graphiques, qui inclut la mise en forme de remplissage et de ligne.|
||[name](/javascript/api/excel/excel.chartseriescollectionloadoptions#name)|Pour chaque élément de la collection: représente le nom d’une série dans un graphique.|
||[pointe](/javascript/api/excel/excel.chartseriescollectionloadoptions#points)|Pour chaque élément de la collection: représente une collection de tous les points de la série.|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[format](/javascript/api/excel/excel.chartseriesdata#format)|Représente le format d’une série de graphique, à savoir le format de remplissage et des lignes. En lecture seule.|
||[name](/javascript/api/excel/excel.chartseriesdata#name)|Représente le nom d’une série dans un graphique.|
||[pointe](/javascript/api/excel/excel.chartseriesdata#points)|Représente la collection de tous les points de la série. En lecture seule.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Représente le format de remplissage d’une série du graphique, qui comprend les informations de mise en forme d’arrière-plan. En lecture seule.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Représente le format des lignes. En lecture seule.|
||[Set (propriétés: Excel. ChartSeriesFormat)](/javascript/api/excel/excel.chartseriesformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartSeriesFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartseriesformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartSeriesFormatData](/javascript/api/excel/excel.chartseriesformatdata)|[line](/javascript/api/excel/excel.chartseriesformatdata#line)|Représente le format des lignes. En lecture seule.|
|[ChartSeriesFormatLoadOptions](/javascript/api/excel/excel.chartseriesformatloadoptions)|[$all](/javascript/api/excel/excel.chartseriesformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.chartseriesformatloadoptions#line)|Représente le format des lignes.|
|[ChartSeriesFormatUpdateData](/javascript/api/excel/excel.chartseriesformatupdatedata)|[line](/javascript/api/excel/excel.chartseriesformatupdatedata#line)|Représente le format des lignes.|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[$all](/javascript/api/excel/excel.chartseriesloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartseriesloadoptions#format)|Représente le format d’une série de graphique, à savoir le format de remplissage et des lignes.|
||[name](/javascript/api/excel/excel.chartseriesloadoptions#name)|Représente le nom d’une série dans un graphique.|
||[pointe](/javascript/api/excel/excel.chartseriesloadoptions#points)|Représente la collection de tous les points de la série.|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[format](/javascript/api/excel/excel.chartseriesupdatedata#format)|Représente le format d’une série de graphique, à savoir le format de remplissage et des lignes.|
||[name](/javascript/api/excel/excel.chartseriesupdatedata#name)|Représente le nom d’une série dans un graphique.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|Valeur booléenne indiquant si le titre du graphique recouvre le graphique ou non.|
||[format](/javascript/api/excel/excel.charttitle#format)|Représente le format du titre d’un graphique, à savoir le format de remplissage et de la police. En lecture seule.|
||[Set (propriétés: Excel. ChartTitle)](/javascript/api/excel/excel.charttitle#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartTitleUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.charttitle#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[text](/javascript/api/excel/excel.charttitle#text)|Représente le texte du titre d’un graphique.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Valeur booléenne qui représente la visibilité d’un objet de titre de graphique.|
|[ChartTitleData](/javascript/api/excel/excel.charttitledata)|[format](/javascript/api/excel/excel.charttitledata#format)|Représente le format du titre d’un graphique, à savoir le format de remplissage et de la police. En lecture seule.|
||[overlay](/javascript/api/excel/excel.charttitledata#overlay)|Valeur booléenne indiquant si le titre du graphique recouvre le graphique ou non.|
||[text](/javascript/api/excel/excel.charttitledata#text)|Représente le texte du titre d’un graphique.|
||[visible](/javascript/api/excel/excel.charttitledata#visible)|Valeur booléenne qui représente la visibilité d’un objet de titre de graphique.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan. En lecture seule.|
||[police](/javascript/api/excel/excel.charttitleformat#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) pour un objet. En lecture seule.|
||[Set (propriétés: Excel. ChartTitleFormat)](/javascript/api/excel/excel.charttitleformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartTitleFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.charttitleformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartTitleFormatData](/javascript/api/excel/excel.charttitleformatdata)|[police](/javascript/api/excel/excel.charttitleformatdata#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) pour un objet. En lecture seule.|
|[ChartTitleFormatLoadOptions](/javascript/api/excel/excel.charttitleformatloadoptions)|[$all](/javascript/api/excel/excel.charttitleformatloadoptions#$all)||
||[police](/javascript/api/excel/excel.charttitleformatloadoptions#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) pour un objet.|
|[ChartTitleFormatUpdateData](/javascript/api/excel/excel.charttitleformatupdatedata)|[police](/javascript/api/excel/excel.charttitleformatupdatedata#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) pour un objet.|
|[ChartTitleLoadOptions](/javascript/api/excel/excel.charttitleloadoptions)|[$all](/javascript/api/excel/excel.charttitleloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttitleloadoptions#format)|Représente le format du titre d’un graphique, à savoir le format de remplissage et de la police.|
||[overlay](/javascript/api/excel/excel.charttitleloadoptions#overlay)|Valeur booléenne indiquant si le titre du graphique recouvre le graphique ou non.|
||[text](/javascript/api/excel/excel.charttitleloadoptions#text)|Représente le texte du titre d’un graphique.|
||[visible](/javascript/api/excel/excel.charttitleloadoptions#visible)|Valeur booléenne qui représente la visibilité d’un objet de titre de graphique.|
|[ChartTitleUpdateData](/javascript/api/excel/excel.charttitleupdatedata)|[format](/javascript/api/excel/excel.charttitleupdatedata#format)|Représente le format du titre d’un graphique, à savoir le format de remplissage et de la police.|
||[overlay](/javascript/api/excel/excel.charttitleupdatedata#overlay)|Valeur booléenne indiquant si le titre du graphique recouvre le graphique ou non.|
||[text](/javascript/api/excel/excel.charttitleupdatedata#text)|Représente le texte du titre d’un graphique.|
||[visible](/javascript/api/excel/excel.charttitleupdatedata#visible)|Valeur booléenne qui représente la visibilité d’un objet de titre de graphique.|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[ordonné](/javascript/api/excel/excel.chartupdatedata#axes)|Représente les axes du graphique.|
||[dataLabels](/javascript/api/excel/excel.chartupdatedata#datalabels)|Représente les étiquettes des données sur le graphique.|
||[format](/javascript/api/excel/excel.chartupdatedata#format)|Regroupe les propriétés de format de la zone de graphique.|
||[height](/javascript/api/excel/excel.chartupdatedata#height)|Représente la hauteur, exprimée en points, de l’objet de graphique.|
||[left](/javascript/api/excel/excel.chartupdatedata#left)|Distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[Legend](/javascript/api/excel/excel.chartupdatedata#legend)|Représente la légende du graphique.|
||[name](/javascript/api/excel/excel.chartupdatedata#name)|Représente le nom d’un objet de graphique.|
||[title](/javascript/api/excel/excel.chartupdatedata#title)|Représente le titre du graphique indiqué et comprend le texte, la visibilité, la position et la mise en forme du titre.|
||[top](/javascript/api/excel/excel.chartupdatedata#top)|Représente la distance, en points, entre le bord supérieur de l’objet et la partie supérieure de la ligne 1 (sur une feuille de calcul) ou le haut de la zone de graphique (sur un graphique).|
||[width](/javascript/api/excel/excel.chartupdatedata#width)|Représente la largeur, en points, de l’objet de graphique.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|Renvoie l’objet de plage qui est associé au nom. Renvoie une erreur si le type de l’élément nommé n’est pas une plage.|
||[name](/javascript/api/excel/excel.nameditem#name)|Nom de l’objet. En lecture seule.|
||[type](/javascript/api/excel/excel.nameditem#type)|Indique le type de la valeur renvoyée par la formule du nom. Pour plus d’informations, voir Excel. NamedItemType. En lecture seule.|
||[value](/javascript/api/excel/excel.nameditem#value)|Représente la valeur calculée par la formule du nom. Pour une plage nommée, renvoie l’adresse de la plage. En lecture seule.|
||[Set (propriétés: Excel. NamedItem)](/javascript/api/excel/excel.nameditem#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. NamedItemUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.nameditem#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Indique si l’objet est visible ou non.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|Obtient un objet NamedItem à l’aide de son nom.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[$all](/javascript/api/excel/excel.nameditemcollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.nameditemcollectionloadoptions#name)|Pour chaque élément de la collection: nom de l’objet. En lecture seule.|
||[type](/javascript/api/excel/excel.nameditemcollectionloadoptions#type)|Pour chaque élément de la collection: indique le type de la valeur renvoyée par la formule du nom. Pour plus d’informations, voir Excel. NamedItemType. En lecture seule.|
||[value](/javascript/api/excel/excel.nameditemcollectionloadoptions#value)|Pour chaque élément de la collection: représente la valeur calculée par la formule du nom. Pour une plage nommée, renvoie l’adresse de la plage. En lecture seule.|
||[visible](/javascript/api/excel/excel.nameditemcollectionloadoptions#visible)|Pour chaque élément de la collection: indique si l’objet est visible ou non.|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[name](/javascript/api/excel/excel.nameditemdata#name)|Nom de l’objet. En lecture seule.|
||[type](/javascript/api/excel/excel.nameditemdata#type)|Indique le type de la valeur renvoyée par la formule du nom. Pour plus d’informations, voir Excel. NamedItemType. En lecture seule.|
||[value](/javascript/api/excel/excel.nameditemdata#value)|Représente la valeur calculée par la formule du nom. Pour une plage nommée, renvoie l’adresse de la plage. En lecture seule.|
||[visible](/javascript/api/excel/excel.nameditemdata#visible)|Indique si l’objet est visible ou non.|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[$all](/javascript/api/excel/excel.nameditemloadoptions#$all)||
||[name](/javascript/api/excel/excel.nameditemloadoptions#name)|Nom de l’objet. En lecture seule.|
||[type](/javascript/api/excel/excel.nameditemloadoptions#type)|Indique le type de la valeur renvoyée par la formule du nom. Pour plus d’informations, voir Excel. NamedItemType. En lecture seule.|
||[value](/javascript/api/excel/excel.nameditemloadoptions#value)|Représente la valeur calculée par la formule du nom. Pour une plage nommée, renvoie l’adresse de la plage. En lecture seule.|
||[visible](/javascript/api/excel/excel.nameditemloadoptions#visible)|Indique si l’objet est visible ou non.|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[visible](/javascript/api/excel/excel.nameditemupdatedata#visible)|Indique si l’objet est visible ou non.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.range#clear-applyto-)|Supprime les valeurs et les propriétés de format, de remplissage, de bordure, etc. de la plage.|
||[Effacer (applyTo ? : Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|Supprime les valeurs et les propriétés de format, de remplissage, de bordure, etc. de la plage.|
||[supprimer (Maj: "vers le \| haut" "gauche")](/javascript/api/excel/excel.range#delete-shift-)|Supprime les cellules associées à la plage.|
||[supprimer (Maj: Excel. DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|Supprime les cellules associées à la plage.|
||[formulas](/javascript/api/excel/excel.range#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
||[getBoundingRect (anotherRange: chaîne \| de plage)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|Renvoie le plus petit objet de plage qui englobe les plages données. Par exemple, la valeur GetBoundingRect pour « B2:C5 » et « D10:E15 » est « B2:E15 ».|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut être située en dehors des limites de sa plage parente, tant qu’elle reste dans la grille de la feuille de calcul. L’emplacement de la cellule renvoyée est déterminé à partir de la cellule supérieure gauche de la plage.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|Obtient une colonne contenue dans la plage.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|Obtient un objet qui représente la colonne entière de la plage (par exemple, si la plage actuelle représente les cellules «B4: E11», `getEntireColumn` qu’il s’agit d’une plage qui représente les colonnes «B:E»).|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|Obtient un objet qui représente la ligne entière de la plage (par exemple, si la plage actuelle représente les cellules «B4: E11», `GetEntireRow` qu’il s’agit d’une plage qui représente les lignes «4:11»).|
||[getIntersection (anotherRange: chaîne \| de plage)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données.|
||[getLastCell()](/javascript/api/excel/excel.range#getlastcell--)|Obtient la dernière cellule de la plage. Par exemple, la dernière cellule de la plage « B2:D5 » est « D5 ».|
||[getLastColumn()](/javascript/api/excel/excel.range#getlastcolumn--)|Obtient la dernière colonne de la plage. Par exemple, la dernière colonne de la plage « B2:D5 » est « D2:D5 ».|
||[getLastRow()](/javascript/api/excel/excel.range#getlastrow--)|Obtient la dernière ligne de la plage. Par exemple, la dernière ligne de la plage « B2:D5 » est « B5:D5 ».|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|Obtient un objet qui représente une plage décalée par rapport à la plage spécifiée. Les dimensions de la plage renvoyée correspondent à cette plage. Si la plage obtenue se retrouve en dehors des limites de grille de la feuille de calcul, une erreur est déclenchée.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|Obtient une ligne contenue dans la plage.|
||[Insérer (Maj: "bas" \| "droite")](/javascript/api/excel/excel.range#insert-shift-)|Insère une cellule ou une plage de cellules dans la feuille de calcul à la place d’une plage donnée et décale les autres cellules pour libérer de l’espace. Renvoie un nouvel objet Range dans l’espace vide qui s’est créé.|
||[Insérer (Maj: Excel. InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|Insère une cellule ou une plage de cellules dans la feuille de calcul à la place d’une plage donnée et décale les autres cellules pour libérer de l’espace. Renvoie un nouvel objet Range dans l’espace vide qui s’est créé.|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|Représente le code de format de nombre d’Excel pour la plage donnée.|
||[adresse](/javascript/api/excel/excel.range#address)|Représente la référence de plage dans le style a1. La valeur de l’adresse contiendra la référence de la feuille (par exemple, «Sheet1! A1: B4 "). En lecture seule.|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|Représente la référence de la plage spécifiée dans le langage de l’utilisateur. En lecture seule.|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|Nombre de cellules dans la plage. Cette API renvoie -1 si le nombre de cellules est supérieur à 2^31-1 (2 147 483 647). En lecture seule.|
||[NbreColonnes](/javascript/api/excel/excel.range#columncount)|Représente le nombre total de colonnes dans la plage. En lecture seule.|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|Représente le numéro de colonne de la première cellule de la plage. Avec indice zéro. En lecture seule.|
||[format](/javascript/api/excel/excel.range#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures, l’alignement et d’autres propriétés de la plage. En lecture seule.|
||[Stopp](/javascript/api/excel/excel.range#rowcount)|Renvoie le nombre total de lignes de la plage. En lecture seule.|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|Renvoie le numéro de ligne de la première cellule de la plage. Avec indice zéro. En lecture seule.|
||[text](/javascript/api/excel/excel.range#text)|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|Représente le type de données de chaque cellule. En lecture seule.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|Feuille de calcul contenant la plage. En lecture seule.|
||[select()](/javascript/api/excel/excel.range#select--)|Sélectionne la plage spécifiée dans l’interface utilisateur d’Excel.|
||[Set (propriétés: Excel. Range)](/javascript/api/excel/excel.range#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. RangeUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.range#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[track()](/javascript/api/excel/excel.range#track--)|Effectuer le suivi de l’objet pour l’ajustement automatique en fonction environnant des modifications dans le document. Cet appel est abréviations context.trackedObjects.add(thisObject). Si vous utilisez cet objet au sein de « .sync » appels et en dehors de l’exécution séquentielle d’un lot de « .run » et rencontrez un message d’erreur « InvalidObjectPath » lors de la définition d’une propriété ou appeler une méthode sur l’objet, vous devez ajouter l’objet à l’objet de suivi collection de sites lors de l’objet a été créé.|
||[untrack()](/javascript/api/excel/excel.range#untrack--)|Publication mémoire associée à cet objet si elle a été précédemment suivie. Cet appel est abréviations context.trackedObjects.add(thisObject). Vous rencontrez de nombreux objets suivies ralentit l’application hôte, donc n’oubliez pas de libérer les objets que l'on ajoute, une fois que vous avez terminé à les utiliser. Vous devez appeler « context.sync() » avant la publication de mémoire prend effet.|
||[values](/javascript/api/excel/excel.range#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|Valeur constante qui indique un côté spécifique de la bordure. Pour plus d’informations, voir Excel. BorderIndex. En lecture seule.|
||[Set (propriétés: Excel. RangeBorder)](/javascript/api/excel/excel.rangeborder#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. RangeBorderUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.rangeborder#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[style](/javascript/api/excel/excel.rangeborder#style)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure. Pour plus d’informations, voir Excel. BorderLineStyle.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Spécifie l'épaisseur de la bordure autour d'une plage. Pour plus d’informations, voir Excel. BorderWeight.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem (index: "EdgeTop" \| "EdgeBottom" \| "EdgeLeft" \| "EdgeRight" \| "InsideVertical" \| "InsideHorizontal" \| "DiagonalDown" \| "DiagonalUp")](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Obtient un objet de bordure à l’aide de son nom.|
||[getItem (index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Obtient un objet de bordure à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|Obtient un objet de bordure à l’aide de son indice.|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|Nombre d’objets de bordure de la collection. En lecture seule.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[RangeBorderCollectionLoadOptions](/javascript/api/excel/excel.rangebordercollectionloadoptions)|[$all](/javascript/api/excel/excel.rangebordercollectionloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangebordercollectionloadoptions#color)|Pour chaque élément de la collection: code couleur HTML qui représente la couleur de la ligne de bordure, de la #RRGGBB de formulaire (par exemple, «FFA500») ou sous forme de couleur HTML nommée (par exemple, «orange»).|
||[sideIndex](/javascript/api/excel/excel.rangebordercollectionloadoptions#sideindex)|Pour chaque élément de la collection: valeur de constante qui indique le côté spécifique de la bordure. Pour plus d’informations, voir Excel. BorderIndex. En lecture seule.|
||[style](/javascript/api/excel/excel.rangebordercollectionloadoptions#style)|Pour chaque élément de la collection: l’une des constantes de style de ligne spécifiant le style de trait de la bordure. Pour plus d’informations, voir Excel. BorderLineStyle.|
||[weight](/javascript/api/excel/excel.rangebordercollectionloadoptions#weight)|Pour chaque élément de la collection: spécifie l’épaisseur de la bordure autour d’une plage. Pour plus d’informations, voir Excel. BorderWeight.|
|[RangeBorderData](/javascript/api/excel/excel.rangeborderdata)|[color](/javascript/api/excel/excel.rangeborderdata#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[sideIndex](/javascript/api/excel/excel.rangeborderdata#sideindex)|Valeur constante qui indique un côté spécifique de la bordure. Pour plus d’informations, voir Excel. BorderIndex. En lecture seule.|
||[style](/javascript/api/excel/excel.rangeborderdata#style)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure. Pour plus d’informations, voir Excel. BorderLineStyle.|
||[weight](/javascript/api/excel/excel.rangeborderdata#weight)|Spécifie l'épaisseur de la bordure autour d'une plage. Pour plus d’informations, voir Excel. BorderWeight.|
|[RangeBorderLoadOptions](/javascript/api/excel/excel.rangeborderloadoptions)|[$all](/javascript/api/excel/excel.rangeborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangeborderloadoptions#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[sideIndex](/javascript/api/excel/excel.rangeborderloadoptions#sideindex)|Valeur constante qui indique un côté spécifique de la bordure. Pour plus d’informations, voir Excel. BorderIndex. En lecture seule.|
||[style](/javascript/api/excel/excel.rangeborderloadoptions#style)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure. Pour plus d’informations, voir Excel. BorderLineStyle.|
||[weight](/javascript/api/excel/excel.rangeborderloadoptions#weight)|Spécifie l'épaisseur de la bordure autour d'une plage. Pour plus d’informations, voir Excel. BorderWeight.|
|[RangeBorderUpdateData](/javascript/api/excel/excel.rangeborderupdatedata)|[color](/javascript/api/excel/excel.rangeborderupdatedata#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[style](/javascript/api/excel/excel.rangeborderupdatedata#style)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure. Pour plus d’informations, voir Excel. BorderLineStyle.|
||[weight](/javascript/api/excel/excel.rangeborderupdatedata#weight)|Spécifie l'épaisseur de la bordure autour d'une plage. Pour plus d’informations, voir Excel. BorderWeight.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[adresse](/javascript/api/excel/excel.rangedata#address)|Représente la référence de plage dans le style a1. La valeur de l’adresse contiendra la référence de la feuille (par exemple, «Sheet1! A1: B4 "). En lecture seule.|
||[addressLocal](/javascript/api/excel/excel.rangedata#addresslocal)|Représente la référence de la plage spécifiée dans le langage de l’utilisateur. En lecture seule.|
||[cellCount](/javascript/api/excel/excel.rangedata#cellcount)|Nombre de cellules dans la plage. Cette API renvoie -1 si le nombre de cellules est supérieur à 2^31-1 (2 147 483 647). En lecture seule.|
||[NbreColonnes](/javascript/api/excel/excel.rangedata#columncount)|Représente le nombre total de colonnes dans la plage. En lecture seule.|
||[columnIndex](/javascript/api/excel/excel.rangedata#columnindex)|Représente le numéro de colonne de la première cellule de la plage. Avec indice zéro. En lecture seule.|
||[format](/javascript/api/excel/excel.rangedata#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures, l’alignement et d’autres propriétés de la plage. En lecture seule.|
||[formulas](/javascript/api/excel/excel.rangedata#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangedata#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
||[numberFormat](/javascript/api/excel/excel.rangedata#numberformat)|Représente le code de format de nombre d’Excel pour la plage donnée.|
||[Stopp](/javascript/api/excel/excel.rangedata#rowcount)|Renvoie le nombre total de lignes de la plage. En lecture seule.|
||[rowIndex](/javascript/api/excel/excel.rangedata#rowindex)|Renvoie le numéro de ligne de la première cellule de la plage. Avec indice zéro. En lecture seule.|
||[text](/javascript/api/excel/excel.rangedata#text)|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|
||[valueTypes](/javascript/api/excel/excel.rangedata#valuetypes)|Représente le type de données de chaque cellule. En lecture seule.|
||[values](/javascript/api/excel/excel.rangedata#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|Réinitialise l’arrière-plan de la plage.|
||[color](/javascript/api/excel/excel.rangefill#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[Set (propriétés: Excel. RangeFill)](/javascript/api/excel/excel.rangefill#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. RangeFillUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.rangefill#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[RangeFillData](/javascript/api/excel/excel.rangefilldata)|[color](/javascript/api/excel/excel.rangefilldata#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
|[RangeFillLoadOptions](/javascript/api/excel/excel.rangefillloadoptions)|[$all](/javascript/api/excel/excel.rangefillloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangefillloadoptions#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
|[RangeFillUpdateData](/javascript/api/excel/excel.rangefillupdatedata)|[color](/javascript/api/excel/excel.rangefillupdatedata#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.rangefont#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.rangefont#name)|Nom de la police (par exemple « Calibri »)|
||[Set (propriétés: Excel. RangeFont)](/javascript/api/excel/excel.rangefont#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. RangeFontUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.rangefont#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[size](/javascript/api/excel/excel.rangefont#size)|Taille de police|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. RangeUnderlineStyle.|
|[RangeFontData](/javascript/api/excel/excel.rangefontdata)|[bold](/javascript/api/excel/excel.rangefontdata#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.rangefontdata#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.rangefontdata#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.rangefontdata#name)|Nom de la police (par exemple « Calibri »)|
||[size](/javascript/api/excel/excel.rangefontdata#size)|Taille de police|
||[underline](/javascript/api/excel/excel.rangefontdata#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. RangeUnderlineStyle.|
|[RangeFontLoadOptions](/javascript/api/excel/excel.rangefontloadoptions)|[$all](/javascript/api/excel/excel.rangefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.rangefontloadoptions#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.rangefontloadoptions#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.rangefontloadoptions#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.rangefontloadoptions#name)|Nom de la police (par exemple « Calibri »)|
||[size](/javascript/api/excel/excel.rangefontloadoptions#size)|Taille de police|
||[underline](/javascript/api/excel/excel.rangefontloadoptions#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. RangeUnderlineStyle.|
|[RangeFontUpdateData](/javascript/api/excel/excel.rangefontupdatedata)|[bold](/javascript/api/excel/excel.rangefontupdatedata#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.rangefontupdatedata#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.rangefontupdatedata#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.rangefontupdatedata#name)|Nom de la police (par exemple « Calibri »)|
||[size](/javascript/api/excel/excel.rangefontupdatedata#size)|Taille de police|
||[underline](/javascript/api/excel/excel.rangefontupdatedata#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. RangeUnderlineStyle.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|Représente l’alignement horizontal de l’objet spécifié. Pour plus d’informations, voir Excel. HorizontalAlignment.|
||[Borders](/javascript/api/excel/excel.rangeformat#borders)|Collection d’objets de bordure qui s’appliquent à l’ensemble de la plage. En lecture seule.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Renvoie l’objet de remplissage défini sur la plage globale. En lecture seule.|
||[police](/javascript/api/excel/excel.rangeformat#font)|Renvoie l’objet de police défini sur l’ensemble de la plage. En lecture seule.|
||[Set (propriétés: Excel. RangeFormat)](/javascript/api/excel/excel.rangeformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. RangeFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.rangeformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|Représente l’alignement vertical de l’objet spécifié. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Indique si Excel renvoie le texte à la ligne dans l’objet. Une valeur null indique que les paramètres de retour à la ligne ne sont pas les mêmes sur l’ensemble de la plage.|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[Borders](/javascript/api/excel/excel.rangeformatdata#borders)|Collection d’objets de bordure qui s’appliquent à l’ensemble de la plage. En lecture seule.|
||[fill](/javascript/api/excel/excel.rangeformatdata#fill)|Renvoie l’objet de remplissage défini sur la plage globale. En lecture seule.|
||[police](/javascript/api/excel/excel.rangeformatdata#font)|Renvoie l’objet de police défini sur l’ensemble de la plage. En lecture seule.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatdata#horizontalalignment)|Représente l’alignement horizontal de l’objet spécifié. Pour plus d’informations, voir Excel. HorizontalAlignment.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatdata#verticalalignment)|Représente l’alignement vertical de l’objet spécifié. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.rangeformatdata#wraptext)|Indique si Excel renvoie le texte à la ligne dans l’objet. Une valeur null indique que les paramètres de retour à la ligne ne sont pas les mêmes sur l’ensemble de la plage.|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[$all](/javascript/api/excel/excel.rangeformatloadoptions#$all)||
||[Borders](/javascript/api/excel/excel.rangeformatloadoptions#borders)|Collection d’objets de bordure qui s’appliquent à l’ensemble de la plage.|
||[fill](/javascript/api/excel/excel.rangeformatloadoptions#fill)|Renvoie l’objet de remplissage défini sur la plage globale.|
||[police](/javascript/api/excel/excel.rangeformatloadoptions#font)|Renvoie l’objet de police défini sur l’ensemble de la plage.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatloadoptions#horizontalalignment)|Représente l’alignement horizontal de l’objet spécifié. Pour plus d’informations, voir Excel. HorizontalAlignment.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatloadoptions#verticalalignment)|Représente l’alignement vertical de l’objet spécifié. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.rangeformatloadoptions#wraptext)|Indique si Excel renvoie le texte à la ligne dans l’objet. Une valeur null indique que les paramètres de retour à la ligne ne sont pas les mêmes sur l’ensemble de la plage.|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[Borders](/javascript/api/excel/excel.rangeformatupdatedata#borders)|Collection d’objets de bordure qui s’appliquent à l’ensemble de la plage.|
||[fill](/javascript/api/excel/excel.rangeformatupdatedata#fill)|Renvoie l’objet de remplissage défini sur la plage globale.|
||[police](/javascript/api/excel/excel.rangeformatupdatedata#font)|Renvoie l’objet de police défini sur l’ensemble de la plage.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatupdatedata#horizontalalignment)|Représente l’alignement horizontal de l’objet spécifié. Pour plus d’informations, voir Excel. HorizontalAlignment.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatupdatedata#verticalalignment)|Représente l’alignement vertical de l’objet spécifié. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.rangeformatupdatedata#wraptext)|Indique si Excel renvoie le texte à la ligne dans l’objet. Une valeur null indique que les paramètres de retour à la ligne ne sont pas les mêmes sur l’ensemble de la plage.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[$all](/javascript/api/excel/excel.rangeloadoptions#$all)||
||[adresse](/javascript/api/excel/excel.rangeloadoptions#address)|Représente la référence de plage dans le style a1. La valeur de l’adresse contiendra la référence de la feuille (par exemple, «Sheet1! A1: B4 "). En lecture seule.|
||[addressLocal](/javascript/api/excel/excel.rangeloadoptions#addresslocal)|Représente la référence de la plage spécifiée dans le langage de l’utilisateur. En lecture seule.|
||[cellCount](/javascript/api/excel/excel.rangeloadoptions#cellcount)|Nombre de cellules dans la plage. Cette API renvoie -1 si le nombre de cellules est supérieur à 2^31-1 (2 147 483 647). En lecture seule.|
||[NbreColonnes](/javascript/api/excel/excel.rangeloadoptions#columncount)|Représente le nombre total de colonnes dans la plage. En lecture seule.|
||[columnIndex](/javascript/api/excel/excel.rangeloadoptions#columnindex)|Représente le numéro de colonne de la première cellule de la plage. Avec indice zéro. En lecture seule.|
||[format](/javascript/api/excel/excel.rangeloadoptions#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures, l’alignement et d’autres propriétés de la plage.|
||[formulas](/javascript/api/excel/excel.rangeloadoptions#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeloadoptions#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
||[numberFormat](/javascript/api/excel/excel.rangeloadoptions#numberformat)|Représente le code de format de nombre d’Excel pour la plage donnée.|
||[Stopp](/javascript/api/excel/excel.rangeloadoptions#rowcount)|Renvoie le nombre total de lignes de la plage. En lecture seule.|
||[rowIndex](/javascript/api/excel/excel.rangeloadoptions#rowindex)|Renvoie le numéro de ligne de la première cellule de la plage. Avec indice zéro. En lecture seule.|
||[text](/javascript/api/excel/excel.rangeloadoptions#text)|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|
||[valueTypes](/javascript/api/excel/excel.rangeloadoptions#valuetypes)|Représente le type de données de chaque cellule. En lecture seule.|
||[values](/javascript/api/excel/excel.rangeloadoptions#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
||[worksheet](/javascript/api/excel/excel.rangeloadoptions#worksheet)|Feuille de calcul contenant la plage.|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[format](/javascript/api/excel/excel.rangeupdatedata#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures, l’alignement et d’autres propriétés de la plage.|
||[formulas](/javascript/api/excel/excel.rangeupdatedata#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeupdatedata#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
||[numberFormat](/javascript/api/excel/excel.rangeupdatedata#numberformat)|Représente le code de format de nombre d’Excel pour la plage donnée.|
||[values](/javascript/api/excel/excel.rangeupdatedata#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|Supprime le tableau.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getdatabodyrange--)|Obtient l’objet de plage associé au corps de données du tableau.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|Obtient l’objet de plage associé à la ligne d’en-tête du tableau.|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|Renvoie l’objet de plage associé à l’intégralité du tableau.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#gettotalrowrange--)|Renvoie l’objet de plage associé à la ligne de total du tableau.|
||[name](/javascript/api/excel/excel.table#name)|Nom du tableau.|
||[colonnes](/javascript/api/excel/excel.table#columns)|Représente une collection de toutes les colonnes du tableau. En lecture seule.|
||[id](/javascript/api/excel/excel.table#id)|Renvoie une valeur qui identifie le tableau dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque le tableau est renommé. En lecture seule.|
||[rows](/javascript/api/excel/excel.table#rows)|Représente une collection de toutes les lignes du tableau. En lecture seule.|
||[Set (propriétés: Excel. table)](/javascript/api/excel/excel.table#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. TableUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.table#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|Indique si la ligne d’en-tête est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne d’en-tête.|
||[ShowTotals,](/javascript/api/excel/excel.table#showtotals)|Indique si la ligne de total est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne de total.|
||[style](/javascript/api/excel/excel.table#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont les suivantes : TableStyleLight1 à TableStyleLight21, TableStyleMedium1 à TableStyleMedium28, TableStyleStyleDark1 à TableStyleStyleDark11. Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[Add (Address: Range \| String, hasHeaders: Boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|Crée un tableau L’adresse de la source ou de l’objet de la plage détermine la feuille de calcul dans laquelle le tableau sera ajouté. Si l’ajout ne peut être effectué (par exemple, parce que l’adresse n’est pas valide, ou parce que le tableau empiéterait sur un autre tableau), un message d’erreur apparaît.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|Obtient un tableau à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|Obtient un tableau en fonction de sa position dans la collection.|
||[count](/javascript/api/excel/excel.tablecollection#count)|Renvoie le nombre de tableaux dans le classeur. En lecture seule.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablecollectionloadoptions#$all)||
||[colonnes](/javascript/api/excel/excel.tablecollectionloadoptions#columns)|Pour chaque élément de la collection: représente une collection de toutes les colonnes du tableau.|
||[id](/javascript/api/excel/excel.tablecollectionloadoptions#id)|Pour chaque élément de la collection: renvoie une valeur qui identifie de manière unique la table dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque le tableau est renommé. En lecture seule.|
||[name](/javascript/api/excel/excel.tablecollectionloadoptions#name)|Pour chaque élément de la collection: nom de la table.|
||[rows](/javascript/api/excel/excel.tablecollectionloadoptions#rows)|Pour chaque élément de la collection: représente une collection de toutes les lignes du tableau.|
||[showHeaders](/javascript/api/excel/excel.tablecollectionloadoptions#showheaders)|Pour chaque élément de la collection: indique si la ligne d’en-tête est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne d’en-tête.|
||[ShowTotals,](/javascript/api/excel/excel.tablecollectionloadoptions#showtotals)|Pour chaque élément de la collection: indique si la ligne total est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne de total.|
||[style](/javascript/api/excel/excel.tablecollectionloadoptions#style)|Pour chaque élément de la collection: valeur de constante qui représente le style de tableau. Les valeurs possibles sont les suivantes : TableStyleLight1 à TableStyleLight21, TableStyleMedium1 à TableStyleMedium28, TableStyleStyleDark1 à TableStyleStyleDark11. Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|Supprime la colonne du tableau.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|Obtient l’objet de plage associé au corps de données de la colonne.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|Obtient l’objet de plage associé à la ligne d’en-tête de la colonne.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|Renvoie l’objet de plage associé à l’intégralité de la colonne.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|Obtient l’objet de plage associé à la ligne de total de la colonne.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Représente le nom de la colonne du tableau.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Renvoie une clé unique qui identifie la colonne du tableau. En lecture seule.|
||[index](/javascript/api/excel/excel.tablecolumn#index)|Renvoie le numéro d’indice de la colonne dans la collection de colonnes du tableau. Avec indice zéro. En lecture seule.|
||[Set (propriétés: Excel. TableColumn)](/javascript/api/excel/excel.tablecolumn#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. TableColumnUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.tablecolumn#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[Add (index?: Number, Values?: Array<Array<\| Boolean \| String Number \|>> \| Boolean \| String Number, Name?: String)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|Ajoute une nouvelle colonne au tableau.|
||[getItem (Key: valeur \| numérique)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|Obtient un objet de colonne par son nom ou son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|Obtient une colonne en fonction de sa position dans la collection.|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|Renvoie le nombre de colonnes du tableau. En lecture seule.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableColumnCollectionLoadOptions](/javascript/api/excel/excel.tablecolumncollectionloadoptions)|[$all](/javascript/api/excel/excel.tablecolumncollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.tablecolumncollectionloadoptions#id)|Pour chaque élément de la collection: renvoie une clé unique qui identifie la colonne dans le tableau. En lecture seule.|
||[index](/javascript/api/excel/excel.tablecolumncollectionloadoptions#index)|Pour chaque élément de la collection: renvoie le numéro d’index de la colonne au sein de la collection Columns du tableau. Avec indice zéro. En lecture seule.|
||[name](/javascript/api/excel/excel.tablecolumncollectionloadoptions#name)|Pour chaque élément de la collection: représente le nom de la colonne de tableau.|
||[values](/javascript/api/excel/excel.tablecolumncollectionloadoptions#values)|Pour chaque élément de la collection: représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableColumnData](/javascript/api/excel/excel.tablecolumndata)|[id](/javascript/api/excel/excel.tablecolumndata#id)|Renvoie une clé unique qui identifie la colonne du tableau. En lecture seule.|
||[index](/javascript/api/excel/excel.tablecolumndata#index)|Renvoie le numéro d’indice de la colonne dans la collection de colonnes du tableau. Avec indice zéro. En lecture seule.|
||[name](/javascript/api/excel/excel.tablecolumndata#name)|Représente le nom de la colonne du tableau.|
||[values](/javascript/api/excel/excel.tablecolumndata#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableColumnLoadOptions](/javascript/api/excel/excel.tablecolumnloadoptions)|[$all](/javascript/api/excel/excel.tablecolumnloadoptions#$all)||
||[id](/javascript/api/excel/excel.tablecolumnloadoptions#id)|Renvoie une clé unique qui identifie la colonne du tableau. En lecture seule.|
||[index](/javascript/api/excel/excel.tablecolumnloadoptions#index)|Renvoie le numéro d’indice de la colonne dans la collection de colonnes du tableau. Avec indice zéro. En lecture seule.|
||[name](/javascript/api/excel/excel.tablecolumnloadoptions#name)|Représente le nom de la colonne du tableau.|
||[values](/javascript/api/excel/excel.tablecolumnloadoptions#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableColumnUpdateData](/javascript/api/excel/excel.tablecolumnupdatedata)|[name](/javascript/api/excel/excel.tablecolumnupdatedata#name)|Représente le nom de la colonne du tableau.|
||[values](/javascript/api/excel/excel.tablecolumnupdatedata#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableData](/javascript/api/excel/excel.tabledata)|[colonnes](/javascript/api/excel/excel.tabledata#columns)|Représente une collection de toutes les colonnes du tableau. En lecture seule.|
||[id](/javascript/api/excel/excel.tabledata#id)|Renvoie une valeur qui identifie le tableau dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque le tableau est renommé. En lecture seule.|
||[name](/javascript/api/excel/excel.tabledata#name)|Nom du tableau.|
||[rows](/javascript/api/excel/excel.tabledata#rows)|Représente une collection de toutes les lignes du tableau. En lecture seule.|
||[showHeaders](/javascript/api/excel/excel.tabledata#showheaders)|Indique si la ligne d’en-tête est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne d’en-tête.|
||[ShowTotals,](/javascript/api/excel/excel.tabledata#showtotals)|Indique si la ligne de total est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne de total.|
||[style](/javascript/api/excel/excel.tabledata#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont les suivantes : TableStyleLight1 à TableStyleLight21, TableStyleMedium1 à TableStyleMedium28, TableStyleStyleDark1 à TableStyleStyleDark11. Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[$all](/javascript/api/excel/excel.tableloadoptions#$all)||
||[colonnes](/javascript/api/excel/excel.tableloadoptions#columns)|Représente une collection de toutes les colonnes du tableau.|
||[id](/javascript/api/excel/excel.tableloadoptions#id)|Renvoie une valeur qui identifie le tableau dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque le tableau est renommé. En lecture seule.|
||[name](/javascript/api/excel/excel.tableloadoptions#name)|Nom du tableau.|
||[rows](/javascript/api/excel/excel.tableloadoptions#rows)|Représente une collection de toutes les lignes du tableau.|
||[showHeaders](/javascript/api/excel/excel.tableloadoptions#showheaders)|Indique si la ligne d’en-tête est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne d’en-tête.|
||[ShowTotals,](/javascript/api/excel/excel.tableloadoptions#showtotals)|Indique si la ligne de total est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne de total.|
||[style](/javascript/api/excel/excel.tableloadoptions#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont les suivantes : TableStyleLight1 à TableStyleLight21, TableStyleMedium1 à TableStyleMedium28, TableStyleStyleDark1 à TableStyleStyleDark11. Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|Supprime la ligne du tableau.|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|Renvoie l’objet de plage associé à la ligne entière.|
||[index](/javascript/api/excel/excel.tablerow#index)|Renvoie le numéro d’indice de la ligne dans la collection de lignes du tableau. Avec indice zéro. En lecture seule.|
||[Set (propriétés: Excel. TableRow)](/javascript/api/excel/excel.tablerow#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. TableRowUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.tablerow#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[values](/javascript/api/excel/excel.tablerow#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[Add (index?: Number, Values?: Array<Array<\| Boolean \| String Number \|>> \| Boolean \| String Number)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|Ajoute une ou plusieurs lignes dans le tableau. L’objet renvoyé sera placé en premier dans les lignes récemment ajoutées.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|Obtient une ligne en fonction de sa position dans la collection.|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|Renvoie le nombre de lignes du tableau. En lecture seule.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableRowCollectionLoadOptions](/javascript/api/excel/excel.tablerowcollectionloadoptions)|[$all](/javascript/api/excel/excel.tablerowcollectionloadoptions#$all)||
||[index](/javascript/api/excel/excel.tablerowcollectionloadoptions#index)|Pour chaque élément de la collection: renvoie le numéro d’index de la ligne dans la collection Rows du tableau. Avec indice zéro. En lecture seule.|
||[values](/javascript/api/excel/excel.tablerowcollectionloadoptions#values)|Pour chaque élément de la collection: représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableRowData](/javascript/api/excel/excel.tablerowdata)|[index](/javascript/api/excel/excel.tablerowdata#index)|Renvoie le numéro d’indice de la ligne dans la collection de lignes du tableau. Avec indice zéro. En lecture seule.|
||[values](/javascript/api/excel/excel.tablerowdata#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableRowLoadOptions](/javascript/api/excel/excel.tablerowloadoptions)|[$all](/javascript/api/excel/excel.tablerowloadoptions#$all)||
||[index](/javascript/api/excel/excel.tablerowloadoptions#index)|Renvoie le numéro d’indice de la ligne dans la collection de lignes du tableau. Avec indice zéro. En lecture seule.|
||[values](/javascript/api/excel/excel.tablerowloadoptions#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableRowUpdateData](/javascript/api/excel/excel.tablerowupdatedata)|[values](/javascript/api/excel/excel.tablerowupdatedata#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableUpdateData](/javascript/api/excel/excel.tableupdatedata)|[name](/javascript/api/excel/excel.tableupdatedata#name)|Nom du tableau.|
||[showHeaders](/javascript/api/excel/excel.tableupdatedata#showheaders)|Indique si la ligne d’en-tête est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne d’en-tête.|
||[ShowTotals,](/javascript/api/excel/excel.tableupdatedata#showtotals)|Indique si la ligne de total est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne de total.|
||[style](/javascript/api/excel/excel.tableupdatedata#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont les suivantes : TableStyleLight1 à TableStyleLight21, TableStyleMedium1 à TableStyleMedium28, TableStyleStyleDark1 à TableStyleStyleDark11. Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange ()](/javascript/api/excel/excel.workbook#getselectedrange--)|Obtient la plage unique actuellement sélectionnée du classeur. Si plusieurs plages sont sélectionnées, cette méthode génère une erreur.|
||[application](/javascript/api/excel/excel.workbook#application)|Représente l’instance de l’application Excel qui contient ce classeur. En lecture seule.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Représente une collection de liaisons appartenant au classeur. En lecture seule.|
||[noms](/javascript/api/excel/excel.workbook#names)|Représente une collection d’éléments nommés portant sur le classeur (appelés plages et constantes). En lecture seule.|
||[emplois](/javascript/api/excel/excel.workbook#tables)|Représente une collection de tableaux associés au classeur. En lecture seule.|
||[feuilles](/javascript/api/excel/excel.workbook#worksheets)|Représente une collection de feuilles de calcul associées au classeur. En lecture seule.|
||[Set (propriétés: Excel. Workbook)](/javascript/api/excel/excel.workbook#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. WorkbookUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.workbook#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[bindings](/javascript/api/excel/excel.workbookdata#bindings)|Représente une collection de liaisons appartenant au classeur. En lecture seule.|
||[noms](/javascript/api/excel/excel.workbookdata#names)|Représente une collection d’éléments nommés portant sur le classeur (appelés plages et constantes). En lecture seule.|
||[emplois](/javascript/api/excel/excel.workbookdata#tables)|Représente une collection de tableaux associés au classeur. En lecture seule.|
||[feuilles](/javascript/api/excel/excel.workbookdata#worksheets)|Représente une collection de feuilles de calcul associées au classeur. En lecture seule.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[$all](/javascript/api/excel/excel.workbookloadoptions#$all)||
||[application](/javascript/api/excel/excel.workbookloadoptions#application)|Représente l’instance de l’application Excel qui contient ce classeur.|
||[bindings](/javascript/api/excel/excel.workbookloadoptions#bindings)|Représente une collection de liaisons appartenant au classeur.|
||[emplois](/javascript/api/excel/excel.workbookloadoptions#tables)|Représente une collection de tableaux associés au classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Active la feuille de calcul dans l’interface utilisateur Excel.|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|Supprime la feuille de calcul du classeur. Notez que si la visibilité de la feuille de calcul est définie sur «VeryHidden», l’opération de suppression échouera avec un GeneralException.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut être située en dehors des limites de sa plage parente, tant qu’elle reste dans la grille de la feuille de calcul.|
||[getRange (Address?: String)](/javascript/api/excel/excel.worksheet#getrange-address-)|Obtient l’objet de plage, représentant un seul bloc de cellules rectangulaires, spécifié par l’adresse ou le nom.|
||[name](/javascript/api/excel/excel.worksheet#name)|Nom complet de la feuille de calcul.|
||[position](/javascript/api/excel/excel.worksheet#position)|Position de la feuille de calcul au sein du classeur (sur une base zéro).|
||[bulles](/javascript/api/excel/excel.worksheet#charts)|Renvoie une collection de graphiques qui font partie de la feuille de calcul. En lecture seule.|
||[id](/javascript/api/excel/excel.worksheet#id)|Renvoie une valeur qui permet d’identifier la feuille de calcul de façon unique dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque la feuille de calcul est renommée ou déplacée. En lecture seule.|
||[emplois](/javascript/api/excel/excel.worksheet#tables)|Collection de tableaux qui font partie de la feuille de calcul. En lecture seule.|
||[Set (propriétés: Excel. Worksheet)](/javascript/api/excel/excel.worksheet#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. WorksheetUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.worksheet#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[excellente](/javascript/api/excel/excel.worksheet#visibility)|Visibilité de la feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[Add (Name?: String)](/javascript/api/excel/excel.worksheetcollection#add-name-)|Ajoute une nouvelle feuille de calcul au classeur. La feuille de calcul est ajoutée à la fin des feuilles de calcul existantes. Si vous souhaitez activer la feuille de calcul nouvellement ajoutée, appelez la méthode .activate() pour cette feuille.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|Obtient la feuille de calcul active du classeur.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[$all](/javascript/api/excel/excel.worksheetcollectionloadoptions#$all)||
||[bulles](/javascript/api/excel/excel.worksheetcollectionloadoptions#charts)|Pour chaque élément de la collection: renvoie la collection de graphiques qui font partie de la feuille de calcul.|
||[id](/javascript/api/excel/excel.worksheetcollectionloadoptions#id)|Pour chaque élément de la collection: renvoie une valeur qui identifie de manière unique la feuille de calcul dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque la feuille de calcul est renommée ou déplacée. En lecture seule.|
||[name](/javascript/api/excel/excel.worksheetcollectionloadoptions#name)|Pour chaque élément de la collection: nom d’affichage de la feuille de calcul.|
||[position](/javascript/api/excel/excel.worksheetcollectionloadoptions#position)|Pour chaque élément de la collection: position de base zéro de la feuille de calcul dans le classeur.|
||[emplois](/javascript/api/excel/excel.worksheetcollectionloadoptions#tables)|Pour chaque élément de la collection: collection de tableaux qui font partie de la feuille de calcul.|
||[excellente](/javascript/api/excel/excel.worksheetcollectionloadoptions#visibility)|Pour chaque élément de la collection: visibilité de la feuille de calcul.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[bulles](/javascript/api/excel/excel.worksheetdata#charts)|Renvoie une collection de graphiques qui font partie de la feuille de calcul. En lecture seule.|
||[id](/javascript/api/excel/excel.worksheetdata#id)|Renvoie une valeur qui permet d’identifier la feuille de calcul de façon unique dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque la feuille de calcul est renommée ou déplacée. En lecture seule.|
||[name](/javascript/api/excel/excel.worksheetdata#name)|Nom complet de la feuille de calcul.|
||[position](/javascript/api/excel/excel.worksheetdata#position)|Position de la feuille de calcul au sein du classeur (sur une base zéro).|
||[emplois](/javascript/api/excel/excel.worksheetdata#tables)|Collection de tableaux qui font partie de la feuille de calcul. En lecture seule.|
||[excellente](/javascript/api/excel/excel.worksheetdata#visibility)|Visibilité de la feuille de calcul.|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[$all](/javascript/api/excel/excel.worksheetloadoptions#$all)||
||[bulles](/javascript/api/excel/excel.worksheetloadoptions#charts)|Renvoie une collection de graphiques qui font partie de la feuille de calcul.|
||[id](/javascript/api/excel/excel.worksheetloadoptions#id)|Renvoie une valeur qui permet d’identifier la feuille de calcul de façon unique dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque la feuille de calcul est renommée ou déplacée. En lecture seule.|
||[name](/javascript/api/excel/excel.worksheetloadoptions#name)|Nom complet de la feuille de calcul.|
||[position](/javascript/api/excel/excel.worksheetloadoptions#position)|Position de la feuille de calcul au sein du classeur (sur une base zéro).|
||[emplois](/javascript/api/excel/excel.worksheetloadoptions#tables)|Collection de tableaux qui font partie de la feuille de calcul.|
||[excellente](/javascript/api/excel/excel.worksheetloadoptions#visibility)|Visibilité de la feuille de calcul.|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[name](/javascript/api/excel/excel.worksheetupdatedata#name)|Nom complet de la feuille de calcul.|
||[position](/javascript/api/excel/excel.worksheetupdatedata#position)|Position de la feuille de calcul au sein du classeur (sur une base zéro).|
||[visibility](/javascript/api/excel/excel.worksheetupdatedata#visibility)|Visibilité de la feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
