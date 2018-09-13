# <a name="build-an-excel-add-in-using-react"></a>Développement d’un complément Excel à l’aide de React

Cet article décrit le processus de création d’un complément Excel à l’aide de React et de l’API JavaScript pour Excel.

## <a name="environment"></a>Environnement

- **Office pour ordinateur de bureau** : Assurez-vous de disposer de la dernière version d'Office. Les commandes du complément nécessitent la version 16.0.6769.0000 ou supérieure (la version **16.0.6868.0000** est conseillée). Apprenez à [Installer la dernière version des applications Office](http://aka.ms/latestoffice). 
 
- **Office Online** : Aucune installation supplémentaire n'est nécessaire. Notez que la prise en charge des commandes au sein d'Office Online pour les comptes professionnels / scolaires est actuellement en préversion.

## <a name="prerequisites"></a>Conditions préalables

- [Node.js](https://nodejs.org)

- Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a>Création de l’application web

1. Créez un dossier sur votre lecteur local et nommez-le **my-addin**. Il s’agit de l’endroit où vous allez créer les fichiers de votre application.

2. Accédez au dossier de votre application.

    ```bash
    cd my-addin
    ```

3. Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément. Exécutez la commande suivante, puis répondez aux invites comme indiqué dans la capture d’écran suivante.

    ```bash
    yo office
    ```

    - **Choisissez un type de projet :** `Office Add-in project using React framework`
    - **Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ? :** `Excel`

    ![Le générateur Yeoman](../images/yo-office-excel-react.png)
    
    Une fois que vous avez terminé avec l'assistant, le générateur crée le projet et installe les composants Node de prise en charge.

4.  Ouvrez **src/components/App.tsx**, recherchez le commentaire « Mettre à jour la couleur de remplissage », puis modifiez la couleur de remplissage de « jaune » à « bleu » avant d'enregistrer le fichier. 

    ```js
    range.format.fill.color = 'blue'

    ```

5. Dans le bloc `return` de la fonction `render` au sein de **src/components/App.tsx**, mettez le `<Herolist>` à jour avec le code ci-dessous, puis enregistrez le fichier. 

    ```js
      <HeroList message='Discover what My Office Add-in can do for you today!' items={this.state.listItems}>
        <p className='ms-font-l'>Choose the button below to set the color of the selected range to blue. <b>Set color</b>.</p>
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
    </HeroList>
    ```

6. Effectuez les étapes décrites dans la rubrique relative à l’[Ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour approuver le certificat pour le système d’exploitation de votre ordinateur de développement.

7. Chargez une version test de votre complément afin qu’il apparaisse dans Excel. Dans le terminal, exécutez la commande suivante : 
    
    ```bash
    npm run sideload
    ```

## <a name="try-it-out"></a>Essayez

1. À partir du terminal, exécutez la commande suivante pour démarrer le serveur dev.

    Windows :
    ```bash
    npm start
    ```

2. |||UNTRANSLATED_CONTENT_START|||In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.|||UNTRANSLATED_CONTENT_END|||

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2b.png)

3. Sélectionnez une plage de cellules dans la feuille de calcul.

4. Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en bleu.

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément Excel à l’aide de React ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.

> [!div class="nextstepaction"]
> [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>Voir aussi

* [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial-create-table.md)
* [Concepts de base de l’API JavaScript pour Excel](../excel/excel-add-ins-core-concepts.md)
* [Exemples de code pour les compléments Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Référence de l’API JavaScript pour Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
