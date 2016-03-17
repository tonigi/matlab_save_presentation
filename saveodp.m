
function saveodp(h,pptxfile,varargin) 
% SAVEODP Create Office Open XML presentations from Matlab. 
%
%   saveodp(h,pptxfile,['-a']) saves figure handle h in a PPTX file.   
%
%   Appends figure as a separate slide at the end of the presentation 
%   if the -a option is present. 
    narginchk(2,3);

    setjavapath;
    import org.odftoolkit.simple.*;
 
    if(nargin==3)
        ppt=PresentationDocument.loadDocument(pptxfile)
    else
        ppt=PresentationDocument.newPresentationDocument();  
      %  ppt.setPageSize(java.awt.Dimension(1200,900));
    end
    
    slide=ppt.newSlide(-1,'No name',SlideLayout.BLANK);
    tn=[tempname '.png'];
    saveas(h,tn,'png');
    
    pictureData = org.apache.commons.io.IOUtils.toByteArray(java.io.FileInputStream(tn));
    idx = ppt.addPicture(pictureData, XSLFPictureData.PICTURE_TYPE_PNG);
    pic = slide.createPicture(idx);

    out = java.io.FileOutputStream(pptxfile);
    ppt.write(out);
    
    delete(tn);

end
 


function setjavapath()
    [here,~,~] = fileparts(mfilename('fullpath'));
    poidir=fullfile(here,'odftoolkit-0.6-incubating/');
    javaaddpath(fullfile(poidir, 'odfdom-java-0.8.9-incubating.jar'));
    javaaddpath(fullfile(poidir, 'simple-odf-0.8-incubating.jar'));
    javaaddpath(fullfile(poidir, 'xslt-runner-1.2.2-incubating.jar'));
    javaaddpath(fullfile(poidir, 'xslt-runner-task-1.2.2-incubating'));
end
