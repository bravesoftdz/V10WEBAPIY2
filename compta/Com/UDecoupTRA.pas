unit UDecoupTRA;

interface
  uses SysUtils, Classes, Dialogs,
  hmsgbox,
  Hctrls, UTOB;


  type
     TabNomExtension = array[1..2] of String;
     TTypeDecoup     = (JOURNAUX,PERIODE);

  type TDecoupTra = class
       private
            procedure RecupLigne(var TRecup, TAutre:TStrings;TypeLigne : string;date1,date2:string);
            procedure RassembleMemJal(var TJal,Tcgn,TCae:TStrings;TAutre:TStrings;adresse:string);
            function  RecupPartie(TAutre:TStrings;typeDebut:string;j:integer):integer;
            procedure DecouperJAL(adresse,date:string);
            procedure DecouperExo(adresse,date:string);

            procedure RecupPeriode(TExo,TAutre,TCae,TCgn:TStrings;adresse:String);

            procedure RecupComptTier(var TCgn,TCae:TStrings);
            procedure RecupComptGnrx(var TEcriture,TCgn:TStrings);
            procedure CreerFichier(chemin:string;TypeDecoup:string;indice:string;nomfich:string;var TRA : TEXTFile);
            function  RecupEcritur(TAutre:TStrings;periode:string;typeRecup:string):TStrings;
       public
            procedure Decouper(adresse : String;TypeDecoup:TTypeDecoup;date:string);
            procedure DecoupAuto(adresse:string);
           
  end;
var TRA              : TextFile;
    TOBEcriture      : TOB;
    FichFrom,FichTo  : TextFile;
    nb               : integer;
implementation

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 01/07/2004
Modifié le ... :   /  /    
Description .. : Decoup le fichier en morceaux sans recup les lignes en tete 
Suite ........ : , juste à la suite
Mots clefs ... : DECOUP AUTO VISU
*****************************************************************}
procedure TDecoupTra.DecoupAuto(adresse:string);
var nomfichier,chemin : string;
    indice,taille,i: integer;
    bloqDecoup        : string;
begin

     nomfichier := ExtractFileName(adresse);
     chemin := ExtractFilePath(adresse);
     indice := 1;
     AssignFile(FichFrom,adresse);
     Reset(FichFrom) ;
     try
        while(not eof(FichFrom))do
        begin
              AssignFile(FichTo,chemin+IntToStr(indice)+nomfichier);
              Rewrite(FichTo);
              i:=0;
              bloqDecoup := '';
              taille := FileSize(FichFrom) div 8;
              taille := (taille div 10 );
              while (( not eof(fichfrom)) and (i<=taille) ) do
              begin
                   Readln(FichFrom,bloqDecoup);
                   Writeln(FichTo,bloqDecoup);
                   i := i+1;
              end;
              CloseFile(FichTo);
              indice := indice + 1;
              PGIINFO('Decoupage d''effectué, veuillez ouvrir un des fichiers créés');
        end;
     finally
           CloseFile(FichFrom);
           PGIInfo('Decoupage impossible');
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : On recup les lignes du type voulu dans un tableau de 
Suite ........ : chaine de caratere TRecup et toutes les autres dans 
Suite ........ : TAutres.
Mots clefs ... : RECUP LIGNES
*****************************************************************}
procedure TDecoupTra.RecupLigne(var TRecup,TAutre:TStrings;TypeLigne : string;date1,date2:string);
var ligne : string;

begin
    Reset(TRA);
    while (not EOF(TRA)) do
    begin
       Readln(TRA,ligne);
       if (Copy(ligne,4,3)<>TypeLigne) then
       begin
          TAutre.Add(ligne);
          nb := nb+1;
       end else
       begin
            if (Copy(ligne,4,3)=TypeLigne)then
            begin
                if ((date1<>'') and (date2<>'')) then
                begin
                     if ((Copy(ligne,14,4)=date1) or (Copy(ligne,14,4)=date2)) then
                        TRecup.Add(ligne);
                        nb := nb+1;

                end else
                begin
                        TRecup.Add(ligne);
                        nb := nb+1;
                end;
            end;
       end;
    end;
    CloseFile(TRA);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Fonction qui ecrit dans le fichier ouvert les parties 
Suite ........ : communes à chaque fichiers découpés
Mots clefs ... : RECUP PARTIE COMMUNE
*****************************************************************}
function  TDecoupTra.RecupPartie(TAutre:TStrings;typeDebut:string;j:integer):integer;
begin
    while((j<=TAutre.Count-1)and (Copy(TAutre.Strings[j],4,3)<>typeDebut)) do
    begin
        Writeln(TRA,TAutre.Strings[j]);
        j := j+1;
    end;
    Result := j;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 11/06/2004
Modifié le ... :   /  /    
Description .. : rassemble les lignes de même type journal dans le même 
Suite ........ : fichier
Mots clefs ... : RASSEMBLE TYPE JOURNAUX
*****************************************************************}
procedure TDecoupTra.RassembleMemJal(var TJal,Tcgn,TCae:TStrings;TAutre:TStrings;adresse:string);
var nomfichier,chemin : string;
    JAL               : string;
    i,j               : integer;
    TEcriture         : TStrings;
begin
    nomfichier := ExtractFileName(adresse);
    chemin := ExtractFilePath(adresse);
    i := 0;
    while i< TJal.Count-1 do
    begin
         JAL := Copy(Tjal.Strings[i],7,3);
         CreerFichier(chemin,'JAL',JAL,nomfichier,TRA);
         RecupPartie(TAutre,'JAL',0);
         Writeln(TRA,Tjal.Strings[i]);
         TJal.Delete(i);
         j := 0;
         while j<=TJal.Count-1 do
         begin
             if (JAL=Copy(TJAL.Strings[j],7,3))then
             begin
                 Writeln(TRA,Tjal.Strings[j]);
                 Tjal.Delete(j);
             end else
                 j := j+1;
         end;
         TEcriture := RecupEcritur(TAutre,JAL,'jal');
         if TEcriture<>nil then
         begin
                RecupComptGnrx(TEcriture,TCgn);
                RecupComptTier(TCgn,TCae);
         end;
         CloseFile(TRA);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 11/06/2004
Modifié le ... :   /  /
Description .. : Decoup les lignes de type Journaux
Mots clefs ... : DECOUP JOURNAUX
*****************************************************************}
procedure TDecoupTra.DecouperJAL(adresse,date:string);
var TJal,TAutre,TCgn,TCae  : TStrings;
begin
    TJal   := TStringList.Create;
    TAutre := TStringList.Create;
    TCgn   := TStringList.Create;
    TCae   := TStringList.Create;

    Assign(TRA,adresse);
    RecupLigne(TCAE,TAutre,'CAE','','');
    Assign(TRA,adresse);
    RecupLigne(TCgn,TAutre,'CGN','','');
    Assign(TRA,adresse);
    RecupLigne(TJal,TAutre,'JAL','','');

    ShowMessage(IntToStr(TAutre.Count));
    RassembleMemJal(TJal,TCGN,TCAe,TAutre,adresse);
    TAutre.Free ;
    TJal.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 11/06/2004
Modifié le ... :   /  /
Description .. : Appel la fonction de decoupage du fichier selon le type
Suite ........ : passé en paramètre
Mots clefs ... : DECOUPER TYPE
*****************************************************************}
procedure TDecoupTra.Decouper(adresse : String;TypeDecoup:TTypeDecoup;date:string); //MARCHE
begin
    AssignFile(TRA,adresse);
    Reset(TRA);
    nb := 0;
    Case TypeDecoup of
       JOURNAUX : DecouperJAL(adresse,date);
       PERIODE  : DecouperExo(adresse,date);
    end;
    PGIInfo('Recup'+IntToStr(nb));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 01/07/2004
Modifié le ... :   /  /    
Description .. : Decoup les fichiers d'apprès les periodes d'exercice
Mots clefs ... : 
*****************************************************************}
procedure TDecoupTra.DecouperExo(adresse,date:string);
var TExo,TAutre,TCAE,TCgn: TStrings;
    date1,date2:string;
begin
     TExo   := TStringList.Create;
     TAutre := TStringList.Create;
     TCgn   := TStringList.Create;
     TCae   := TStringList.Create;
     Assign(TRA,adresse);
     RecupLigne(TCAE,TAutre,'CAE','','');
     Assign(TRA,adresse);
     RecupLigne(TCgn,TAutre,'CGN','','');
     Assign(TRA,adresse);
     if date='Automatique' then
     begin
        date1:='';
        date2:='';
     end else
     begin
         date2:=ReadTokenPipe(date,'/');
         date1:= date;
     end;
     RecupLigne(TExo,TAutre,'EXO',date1,date2);
     RecupPeriode(TExo,TAutre,TCAE,TCgn,adresse);
     TExo.Free;
     Tcae.Free;
     TCgn.Free;
     TAutre.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 01/07/2004
Modifié le ... :   /  /    
Description .. : Recup les periodes correspondant aux exercices
Mots clefs ... : RECUP PERIODE
*****************************************************************}
procedure TDecoupTra.RecupPeriode(TExo,TAutre,TCae,TCgn:TStrings;adresse:String);
var periode,chemin,nomfichier : string;
    i                         : integer;
    TEcriture                 : TStrings;
begin
    nomfichier := ExtractFileName(adresse);
    chemin     := ExtractFilePath(adresse);
    TEcriture  := TStringList.Create;
    i := 0;
    while i<=TExo.Count-1 do
    begin
       periode := Copy(TExo.Strings[i],14,4);
       if ((Trim(periode)<>'') and (length(periode)<=8)) then
       begin
            CreerFichier(chemin,'Periode',periode,nomfichier,TRA);
            RecupPartie(TAutre,'EXO',0);
            Writeln(TRA,TExo.Strings[i]);
            TEcriture := RecupEcritur(TAutre,periode,'exo');
            if TEcriture<>nil then
            begin
                RecupComptGnrx(TEcriture,TCgn);
                RecupComptTier(TCgn,TCae);
            end;
        end;
       i := i+1;
   end;
   CloseFile(TRA);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Recupere les ecritures concernée par la periode de 
Suite ........ : decoupage
Mots clefs ... : RECUP ECRITURE PERIODE
*****************************************************************}
function TDecoupTra.RecupEcritur(TAutre:TStrings;periode:string;typeRecup:string):TStrings;
var TEcriture : TStrings;
    i,trouv   : integer;
    borne1,borne2 : integer;
    typetra : string;
begin
    i := 0;
    trouv := 0;
    borne1 := 0;
    borne2 := 0;

    if typeRecup = 'jal' then
    begin
        borne1 := 1;
        borne2 := 3;
    end else
    begin
         if typeRecup ='exo' then
         begin
             borne1 := 8;
             borne2 := 4;
         end;
    end;
    TEcriture := TStringList.Create;

       while ((i<=TAutre.Count-1) and (Copy(TAutre.Strings[i],1,3)='***')) do
       begin
            i := i+1;
       end;
       while(i<=TAutre.Count-1)do
       begin
            typetra:= Trim(Copy(TAutre.Strings[i],borne1,borne2));
            if(typetra<>'')then
            begin

                 if(typetra = Trim(periode)) then
                 begin
                     Writeln(TRA,TAutre.Strings[i]);
                     TEcriture.Add(TAutre.Strings[i]);
                     trouv := 1+trouv;
                 end;
            end;
            i := i+1;
        end;
        if trouv>0 then
        begin
           Result := TEcriture;
        end else
           Result := nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Recup les comptes Generaux correspondant aux Ecriture 
Suite ........ : de la periode concernée par le decoupage
Mots clefs ... : RECUP COMPT GNRX PERIODE
*****************************************************************}
procedure TDecoupTra.RecupComptGnrx(var TEcriture,TCgn:TStrings);
var i,j: integer;
begin
    i := 0;
    while(i<=TEcriture.Count-1) do
    begin
         for j:=0 to TCgn.Count-1 do
         begin
             if(Trim(Copy(TEcriture.Strings[i],14,31))=Trim(Copy(TCgn.Strings[j],7,24)))then
                if Trim(Copy(TEcriture.Strings[i],14,31))<>'' then
                   Writeln(TRA,TCgn.Strings[j]);
         end;
         i := i+1;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Recup les comptes Tiers contenu dans les 
Suite ........ : compte sgeneraux , correspondant aux Ecriture 
Suite ........ : de la periode concernée par le decoupage
Mots clefs ... : RECUP COMPT TIERS PERIODE
*****************************************************************}
procedure TDecoupTra.RecupComptTier(var TCgn,TCae:TStrings);
var i,j : integer;
begin
    i := 0;
    while(i<=TCgn.Count-1)do
    begin
         for j:=0 to TCae.Count-1 do
         begin
              if(Trim(Copy(TCAE.Strings[j],63,80))=Trim(Copy(TCgn.Strings[i],7,24)))then
                 Writeln(TRA,Tcae.Strings[j]);
         end;
         i := i+1;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Ouvre un fichier en creation
Mots clefs ... : CREER FICHIER
*****************************************************************}
procedure TDecoupTra.CreerFichier(chemin:string;TypeDecoup:string;indice:string;nomfich:string;var TRA : TEXTFile);
begin
     AssignFile(TRA,chemin+TypeDecoup+indice+'_'+nomfich);
     Rewrite(TRA);
end;

end.
