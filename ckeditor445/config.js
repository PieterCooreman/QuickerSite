/**
 * @license Copyright (c) 2003-2013, CKSource - Frederico Knabben. All rights reserved.
 * For licensing, see LICENSE.md or http://ckeditor.com/license
 */

CKEDITOR.editorConfig = function( config ) {
	// Define changes to default configuration here. For example:
	// config.language = 'fr';
	// config.uiColor = '#AADC6E';
	config.allowedContent = true;
	config.resize_dir = 'both';
};

//aanpassingen nav QUICKERSITE
CKEDITOR.on('instanceReady', function(ev)
    {
        var tags = ['p', 'ol', 'ul', 'li']; // etc.

        for (var key in tags) {
            ev.editor.dataProcessor.writer.setRules(tags[key],
                {
                    indent : false,
                    breakBeforeOpen : true,
                    breakAfterOpen : false,
                    breakBeforeClose : false,
                    breakAfterClose : true
                });
        }
    });
	

CKEDITOR.config.toolbar_siteBuilder = 
[
	['Save','Preview','-','Templates'],
	['Cut','Copy','Paste','-','Print'],
	['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat','-','Source','Maximize','ShowBlocks','-','About'],
	'/',
	['Bold','Italic','Underline','Strike','-','Subscript','Superscript'],
	['NumberedList','BulletedList','-','Outdent','Indent','Blockquote','CreateDiv'],
	['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
	['Link','Unlink','Anchor'],
	['Image','Flash','Table','HorizontalRule','Smiley','SpecialChar'],
	'/',
	['Format','Font','FontSize'],
	['TextColor','BGColor']	
] ;

CKEDITOR.config.toolbar_siteBuilderBanner = 
[
	['Save','Bold','Italic','Underline','Link'],
	['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
	['Image','Flash','HorizontalRule','Smiley','NumberedList','BulletedList'],
	['FontSize','TextColor','BGColor','Source']
] ;

CKEDITOR.config.toolbar_siteBuilderMail = 
[
	['Paste','Bold','Italic','Underline'],
	['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
	['Link','Image','Smiley'],
	['NumberedList','BulletedList'],
	['TextColor','BGColor','Source']
] ;

CKEDITOR.config.toolbar_siteBuilderMailNoSource = 
[
	['Paste','Bold','Italic','Underline'],
	['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
	['Link','Image','Smiley'],
	['NumberedList','BulletedList'],
	['TextColor','BGColor']
] ;

CKEDITOR.config.toolbar_siteBuilderMailSource = 
[
	['Paste','Bold','Italic','Underline'],
	['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
	['Link','Image','Smiley'],
	['NumberedList','BulletedList'],
	['TextColor','BGColor','Source']
] ;
//einde aanpassingen