unit CreationMaq;

interface
uses
  Windows, Messages, Graphics, Dialogs, Sysutils, Gridsynth,Critedt,
  fe_main;

Procedure CreationMaquette (typemaq : integer);

implementation
Procedure CreationMaquette (typemaq : integer);
begin
  case typemaq of
  1 : // compte de résultat
  begin
       AGLLanceFiche('CP','CREATMODIFMAQ','','','CR');
  end;
  2 : // solde intermédiaire de gestion
       AGLLanceFiche('CP','CREATMODIFMAQ','','','SIG');
  3 : // bilan
       AGLLanceFiche('CP','CREATMODIFMAQ','','','BIL');
  end;
end;


end.
 