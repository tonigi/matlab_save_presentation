
function savepptx(h,pptxfile,varargin)
% SAVEPPTX Create Office Open XML presentations from Matlab.
%
%   savepptx(h,pptxfile,['-a']) saves figure handle h in a PPTX file.
%
%   Appends figure as a separate slide at the end of the presentation
%   if the -a option is present.
narginchk(2,3);

setjavapath; 

% Eval nonsense is an attempt to workaround this bug http://www.mathworks.com/matlabcentral/answers/48233-how-come-javaaddpath-import-do-not-work-in-a-script-function
eval('import org.apache.poi.xslf.usermodel.*');

if(nargin==3)
    ppt=XMLSlideShow(java.io.FileInputStream(pptxfile));
else
    ppt=XMLSlideShow();
    ppt.setPageSize(java.awt.Dimension(1200,900));
end

slide=ppt.createSlide();
tn=[tempname '.png'];
saveas(h,tn,'png');

pictureData = org.apache.commons.io.IOUtils.toByteArray(java.io.FileInputStream(tn));
idx = ppt.addPicture(pictureData, XSLFPictureData.PICTURE_TYPE_PNG);
pic = slide.createPicture(idx);

out = java.io.FileOutputStream(pptxfile);
ppt.write(out);
out.close();

delete(tn);

end


function setjavapath()

% for some reason global and persistent do not work
% Run only once
if(any(cell2mat(regexp(javaclasspath,'poi-ooxml'))))
    return
end

[here,~,~] = fileparts(mfilename('fullpath'));
poidir=fullfile(here,'poi-3.10-FINAL/');
for i={'ooxml-lib/xmlbeans-2.3.0.jar'   ...
        'ooxml-lib/dom4j-1.6.1.jar'   ...
        'poi-3.10-FINAL-20140208.jar'   ...
        'poi-ooxml-3.10-FINAL-20140208.jar'  ...
        'poi-ooxml-schemas-3.10-FINAL-20140208.jar'}
    jar=fullfile(poidir,i{:});
    javaaddpath(jar);
end

end
