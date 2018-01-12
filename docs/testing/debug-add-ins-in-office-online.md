
# <a name="debug-add-ins-in-office-online"></a>Débogage de compléments dans Office Online


Vous pouvez créer et déboguer des compléments sur un ordinateur n’exécutant pas Windows, ou le client de bureau Office 2013 ou Office 2016 (par exemple, si vous développez sur un Mac). Cet article décrit la procédure d’utilisation d’Office Online dans le but de tester et de déboguer vos compléments. 

Mise en route :


- Créez un compte de développeur Office 365 (si vous n’en avez pas) ou accédez à un site SharePoint.
    
     >**Remarque**  Pour vous inscrire à un compte de développeur Office 365 gratuit, participez à notre [programme de développement Office 365](https://dev.office.com/devprogram).
     
- Configurez un catalogue de compléments sur Office 365 (SharePoint Online). Un catalogue de compléments est une collection de sites dédiée dans SharePoint Online qui héberge des bibliothèques de documents pour des compléments Office. Si vous disposez de votre propre site SharePoint, vous pouvez configurer une bibliothèque de document de catalogue de compléments. Pour plus d’informations, voir [Publier des compléments de contenu et du volet Office dans un catalogue de compléments sur SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a>Débogage de compléments à partir d’Excel Online ou de Word Online

Pour déboguer votre complément à l’aide d’Office Online, procédez comme suit :


1. Déployez votre complément vers un serveur prenant en charge le protocole SSL.
    
     >**Remarque :**  Nous vous recommandons d’utiliser le [générateur Yeoman](https://github.com/OfficeDev/generator-office) pour créer et héberger votre complément.
     
2. Dans le [fichier manifeste de votre complément](../../docs/overview/add-in-manifests.md), mettez à jour la valeur de l’élément **SourceLocation** afin d’inclure un URI absolu, plutôt que relatif. Par exemple :
    
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. Téléchargez le manifeste dans la bibliothèque de compléments Office du catalogue de compléments sur SharePoint.
    
4. Lancez Excel Online ou Word Online à partir du lanceur d’applications dans Office 365, puis ouvrez un nouveau document.
    
5. Sur l’onglet Insérer, sélectionnez  **Mes compléments** ou **Compléments Office** pour insérer votre complément et le tester dans l’application.
    
6. Utilisez l’outil de débogage de votre navigateur préféré pour déboguer votre complément.
    
    Voici certains problèmes que vous pouvez rencontrer lorsque vous effectuez des opérations de débogage :
    
  - Certaines erreurs JavaScript peuvent provenir d’Office Online.
    
  - Le navigateur peut afficher une erreur liée à un certificat non valide que vous devrez contourner.
    
  - Si vous définissez des points d’arrêt dans votre code, Office Online peut générer une erreur indiquant qu’il ne peut pas effectuer d’enregistrement.
    

## <a name="additional-resources"></a>Ressources supplémentaires


- [Meilleures pratiques en matière de développement de compléments Office](../overview/add-in-development-best-practices.md)
    
- [Stratégies de validation pour les applications et les compléments envoyés à l’Office Store (version 1.9)](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)
    
- [Création d’applications et de compléments efficaces pour l’Office Store](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)
    
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
    
