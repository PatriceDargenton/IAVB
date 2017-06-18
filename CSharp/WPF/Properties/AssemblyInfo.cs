﻿
using System;
using System.Reflection;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;

// Les informations générales relatives à un assembly dépendent de 
// l'ensemble d'attributs suivant. Changez les valeurs de ces attributs pour modifier les informations
// associées à un assembly.
[assembly: AssemblyTitle("IAVB3 - C# WPF")]
[assembly: AssemblyDescription(
    "Intelligence Artificielle Vraiment Basique (C# WPF) - Par Patrice Dargenton." + 
    " Logiciel original : publication de Philippe LARVET dans MICRO-SYSTEMES en Déc. 1984")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("ORS Production")]
[assembly: AssemblyProduct("IAVB3 - C# WPF")]
[assembly: AssemblyCopyright("Copyright © ORS Production 2017")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// L'affectation de la valeur false à ComVisible rend les types invisibles dans cet assembly 
// aux composants COM.  Si vous devez accéder à un type dans cet assembly à partir de 
// COM, affectez la valeur true à l'attribut ComVisible sur ce type.
[assembly: ComVisible(false)]

[assembly: CLSCompliant(true)] // CA1014 : Marquer les assemblys avec CLSCompliantAttribute

//Pour commencer à générer des applications localisables, définissez 
//<UICulture>CultureYouAreCodingWith</UICulture> dans votre fichier .csproj
//dans <PropertyGroup>.  Par exemple, si vous utilisez le français
//dans vos fichiers sources, définissez <UICulture> à fr-FR.  Puis, supprimez les marques de commentaire de
//l'attribut NeutralResourceLanguage ci-dessous.  Mettez à jour "fr-FR" dans
//la ligne ci-après pour qu'elle corresponde au paramètre UICulture du fichier projet.

//[assembly: NeutralResourcesLanguage("en-US", UltimateResourceFallbackLocation.Satellite)]

[assembly: ThemeInfo(
    ResourceDictionaryLocation.None, //où se trouvent les dictionnaires de ressources spécifiques à un thème
    //(utilisé si une ressource est introuvable dans la page, 
    // ou dictionnaires de ressources de l'application)
    ResourceDictionaryLocation.SourceAssembly //où se trouve le dictionnaire de ressources générique
    //(utilisé si une ressource est introuvable dans la page, 
    // dans l'application ou dans l'un des dictionnaires de ressources spécifiques à un thème)
)]

[assembly: AssemblyVersion("3.1.2.*")]
//[assembly: AssemblyFileVersion("1.0.0.0")]
