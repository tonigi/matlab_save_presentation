
function saveodp(h,odpfile,varargin) 
% SAVEODP Create OpenOffice/LibreOffice presentations from Matlab. 
%
%   saveodp(h,odpfile,['-a']) saves figure h in odpfile.   
%
% Appends figure as a separate slide at the end of the presentation 
% if the -a option is present. 

    narginchk(2,3);
    
    exe='img2odp.pl';

    opt_a='';
    if nargin==3 
        opt_a='-a';
    end

    pngfile=[tempname '.png'];
    saveas(h,pngfile,'png');
    
    cmd=sprintf('perl %s %s -o %s %s',exe,opt_a,odpfile,pngfile);
    [e,r]=system(cmd);
    
    if e ~= 0 
        error(r);
    end
   
    delete(pngfile);
end

