
function varargout = imgdisp(varargin)
% IMGDISP MATLAB code for imgdisp.fig
%      IMGDISP, by itself, creates a new IMGDISP or raises the existing
%      singleton*.
%
%      H = IMGDISP returns the handle to a new IMGDISP or the handle to
%      the existing singleton*.
%
%      IMGDISP('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in IMGDISP.M with the given input arguments.
%
%      IMGDISP('Property','Value',...) creates a new IMGDISP or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before imgdisp_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to imgdisp_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help imgdisp

% Last Modified by GUIDE v2.5 19-May-2021 20:10:54

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @imgdisp_OpeningFcn, ...
                   'gui_OutputFcn',  @imgdisp_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before imgdisp is made visible.
function imgdisp_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to imgdisp (see VARARGIN)

% Choose default command line output for imgdisp
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes imgdisp wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = imgdisp_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

[rawname,rawpath]=uigetfile({'*.*'},'Select image data');
fullname=[rawpath,rawname];
im=imread(fullname);
handles.blankimage=im;
guidata(hObject, handles);

function edit1_Callback(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit1 as text
%        str2double(get(hObject,'String')) returns contents of edit1 as a double


% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
[handles.filename,filepath1]=uigetfile({'*.*','All Files'},...
  'Select Data File 1');
  cd(filepath1);
  
  

guidata(hObject, handles);
    
 


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
 [num,txt] = xlsread(handles.filename)%read xlsfile,text and numbers are stored diffrently

len=length(txt)%it tells how many certificates are to be generated
for i=1:len
    for j=2:2
        name(i,j)=txt(i,j);
    end
end
for i=1:len
    for j= 3:3
       internship_topic(i,j)=txt(i,j);
    end
end
for i=1:len
    for j= 4:4
       date(i,j)=txt(i,j);
    end
end
for i=1:len
    for j= 5:5
       department(i,j)=txt(i,j);
    end
end
for i=2:len
    combined_text=[name(i,2) internship_topic(i,3) department(i,5)];
    date_text=[date(i,4)];
    position=[1050 630; 800 780;1265 860];
    position2=[270 870];
    RGB = insertText(handles.blankimage,position,combined_text,'FontSize',45,'BoxOpacity',0,'Font','Times New Roman Bold');
    RGB2 = insertText(RGB,position2,date_text,'FontSize',40,'BoxOpacity',0,'Font','Times New Roman Bold');
    imshow(RGB2);
    y=i-1 
    filename=['Certificate_Topic_' num2str(y)]
    lastf=[filename '.jpeg']
    F = getframe(handles.axes1);
    Image = frame2im(F);
    imwrite(Image, lastf)
end
