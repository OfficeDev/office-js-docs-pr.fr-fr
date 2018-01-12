
#<a name="use-office-ui-fabric-261-in-office-add-ins"></a>Utiliser Office UI Fabric 2.6.1 dans des compléments Office

Si vous créez un complément Office, nous vous encourageons à utiliser [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) pour mettre au point l’expérience utilisateur. La procédure suivante présente les opérations de base pour l’utilisation de Fabric.  

> **Remarque :** pour plus d’informations sur Office UI Fabric JS, consultez l’article relatif à l’[utilisation d’Office UI Fabric dans les compléments Office](https://dev.office.com/docs/add-ins/design/using-office-ui-fabric-js).

##<a name="1-set-up-fabric"></a>1. Configuration de Fabric
Ajoutez les lignes suivantes à votre code HTML dans la section d’en-tête pour référencer la structure à partir du réseau de diffusion de contenu (CDN).

     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">


##<a name="2-use-fabric-icons-and-fonts"></a>2. Utiliser les polices et les icônes de la structure
Les icônes sont très simples à utiliser. Il vous suffit d’utiliser un élément « i » et de référencer les classes appropriées. Vous pouvez contrôler la taille de l’icône en modifiant la taille de police.

    <i class="ms-Icon ms-Icon--group" style="font-size:xx-large" aria-hidden="true"></i>


##<a name="3-use-styles-for-simple-components"></a>3. Utiliser des styles pour les composants simples
Fabric comporte des styles pour différents éléments de l’interface utilisateur, tels que des boutons et des cases à cocher. Il vous suffit de référencer les classes appropriées pour ajouter le style correspondant, comme illustré dans l’exemple suivant.

    <button class="ms-Button" id="get-data-from-selection">
    <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
    <span class="ms-Button-label">Get Data from selection</span>
    <span class="ms-Button-description">Get Data from the document selection</span>
    </button>

##<a name="4-use-components-with-sample-behavior"></a>4. Utiliser des composants avec des exemples de comportement
Fabric inclut certains composants qui prennent en charge les comportements (par exemple, ce qu’il se passe lorsque l’utilisateur clique sur un bouton de la souris). Pour vous aider, la **version 2.6.1 de la structure** inclut des **exemples de code** sous la forme de plug-ins d’interface utilisateur JQuery que vous pouvez utiliser. Vous pouvez également utiliser n’importe quelle autre infrastructure pour tout faire fonctionner. Si vous choisissez d’utiliser les exemples fournis, notez que ce code n’est pas distribué par le CDN. Vous devrez donc le télécharger à partir de la **version 2.6.1** du [projet GitHub de la structure](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1), le référencer, puis l’initialiser au sein de votre code. 

Par exemple, pour utiliser le composant SearchBox :

1. Téléchargez le composant SearchBox à partir de [GitHub](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1/src/components/SearchBox).
2. Ajoutez la référence suivante à votre code : `<script src="SearchBox/Jquery.SearchBox.js"></script>`
3. Initialisez le composant en vous assurant que la ligne suivante est exécutée lors du chargement de votre page : `$(".ms-SearchBox").SearchBox();`. Nous vous conseillons d’inclure cette ligne dans le bloc `Office.Initialize` de votre complément.     

**Remarque :** si vous ne comptez pas utiliser tous les composants de la structure, vous pouvez réduire le volume de ressources téléchargées en hébergeant les fichiers CSS individuels pour chaque composant. Vous pouvez obtenir les fichiers CSS dans les dossiers des composants du [référentiel GitHub de la version 2.6.1 de la structure](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1). 


##<a name="next-steps"></a>Étapes suivantes
Si vous cherchez des exemples complets montrant comment utiliser Office UI Fabric, nous avons tout prévu. Reportez-vous aux [exemples d’éléments d’Office UI Fabric pour les compléments Office](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample). Vous pouvez également explorer le site web interactif d’[Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric).

