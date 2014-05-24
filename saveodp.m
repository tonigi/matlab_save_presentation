
function saveodp(h,odpfile,varargin) 
% SAVEODP Create OpenOffice/LibreOffice presentations from Matlab. 
%
%   saveodp(h,odpfile,['-a']) saves figure h in odpfile.   
%
% Appends figure as a separate slide at the end of the presentation 
% if the -a option is present. 
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
    do=xmlread(cont);
    % do=xmlread('test/content.xml')
    op=do.getElementsByTagName('office:presentation').item(0);
    
    slide0=op.getChildNodes.item(0);
    
    slideN=slide0.cloneNode(true);
    slideN.getFirstChild.getFirstChild.setAttribute('xlink:href',pngfile);
    
    op.appendChild(slideN);
    xmlwrite(cont,do);
end