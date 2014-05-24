
function saveodp(h,odpfile,varargin) 
% SAVEODP Create OpenOffice/LibreOffice presentations from Matlab. 
%
%   saveodp(h,odpfile,['-a']) saves figure h in odpfile.   
%
%   Appends figure as a separate slide at the end of the presentation 
%   if the -a option is present. The presentation is based on the file
%   'template.odp' in the same directory as saveodp.m (the first slide is
%   cloned).
    narginchk(2,3);
    [mydir,~,~] = fileparts(mfilename('fullpath'));
    odptemplate=fullfile(mydir,'template.odp');
    
    if(nargin==3)
        infile=odpfile;
    else
        infile=odptemplate;
    end

    % Unzip in a temp dir
    td=unodp(infile);

    % Write PNG in there
    [~,basename,~]=fileparts(td);
    pngfile=fullfile('Pictures',[basename '.png']);
    saveas(h,fullfile(td,pngfile),'png');

    % Fixup content
    addslide(td,pngfile);
   
    % Re-pack and delete
    reodp(odpfile,td);
end

function td=unodp(fn) 
    td=tempname;
    unzip(fn,td);
end

% Deletes odpdir
function reodp(odp,odpdir)
    zip(odp,'*',odpdir);
    movefile([odp '.zip'],odp);
    rmdir(odpdir,'s');
end

function addslide(td,pngfile)
    cont=fullfile(td,'content.xml');
    document=xmlread(cont);
    % do=xmlread('test/content.xml')
    presentation=document.getElementsByTagName('office:presentation').item(0);
    
    % clone the 1st slide, assumed to be the 1st child of presentation
    page1=document.getElementsByTagName('draw:page').item(0);
    newPage=page1.cloneNode(true);
    newImage=newPage.getElementsByTagName('draw:image').item(0);
    newImage.setAttribute('xlink:href',pngfile);
    %        ^ draw:frame  ^ draw:image
    
    presentation.appendChild(newPage);
    xmlwrite(cont,document);
end

% plot(rand(10))
% saveodp(gcf,'x.odp')
% !unzip x.odp content.xml
