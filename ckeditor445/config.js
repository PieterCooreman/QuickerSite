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
	
CKEDITOR.config.templates_replaceContent = false;

/*
 * Full CKEditor 4 Style Set for Modern Layouts
 * Mimicking Bootstrap 5.3 and Modern Design Trends
 */

CKEDITOR.stylesSet.add('default', [
    /* --- SECTION: ALERTS (BOOTSTRAP 5.3) --- */
	{
		name: 'Clear Style (Normal Text)',
		element: 'p',
		attributes: { 'class': '', 'style': '' }
	},
    {
        name: 'Alert: Green',
        element: 'div',
        attributes: {             
            'style': 'padding: 1rem; margin-bottom: 1rem; border: 1px solid #badbcc; border-radius: 0.375rem; background-color: #d1e7dd; color: #0a3622;'
        }
    },
    {
        name: 'Alert: Red',
        element: 'div',
        attributes: {            
            'style': 'padding: 1rem; margin-bottom: 1rem; border: 1px solid #f1aeb5; border-radius: 0.375rem; background-color: #f8d7da; color: #58151c;'
        }
    },
    {
        name: 'Alert: Blue',
        element: 'div',
        attributes: {             
            'style': 'padding: 1rem; margin-bottom: 1rem; border: 1px solid #b6effb; border-radius: 0.375rem; background-color: #cff4fc; color: #055160;'
        }
    },
    {
        name: 'Alert: Yellow',
        element: 'div',
        attributes: {             
            'style': 'padding: 1rem; margin-bottom: 1rem; border: 1px solid #ffecb5; border-radius: 0.375rem; background-color: #fff3cd; color: #664d03;'
        }
    },

    /* --- SECTION: BUTTONS (LINK CONVERSION) --- */
    {
        name: 'Button: Blue',
        element: 'a',
        attributes: { 
           'style': 'display: inline-block; font-weight: 400; line-height: 1.5; color: #fff; background-color: #0D6EFD; padding: 10px 20px; text-decoration: none; border-radius: 0.375rem; border: 1px solid #0D6EFD;'
        }
    },
    {
        name: 'Button: Green',
        element: 'a',
        attributes: {            
            'style': 'display: inline-block; font-weight: 400; line-height: 1.5; color: #fff; background-color: #198754; padding: 10px 20px; text-decoration: none; border-radius: 0.375rem; border: 1px solid #198754;'
        }
    },
	{
        name: 'Button: Yellow',
        element: 'a',
        attributes: {             
            'style': 'display: inline-block; font-weight: 400; line-height: 1.5; color: #664d03; background-color: #FFC107; padding: 10px 20px; text-decoration: none; border-radius: 0.375rem; border: 1px solid #FFC107;'
        }
    },
	{
        name: 'Button: Cyan',
        element: 'a',
        attributes: {             
            'style': 'display: inline-block; font-weight: 400; line-height: 1.5; color: #FFF; background-color: #0DCAF0; padding: 10px 20px; text-decoration: none; border-radius: 0.375rem; border: 1px solid #0DCAF0;'
        }
    },
	{
        name: 'Button: Dark',
        element: 'a',
        attributes: {             
            'style': 'display: inline-block; font-weight: 400; line-height: 1.5; background-color: #212529; color:#EEE; padding: 10px 20px; text-decoration: none; border-radius: 0.375rem; border: 1px solid #212529;'
        }
    },
	{
        name: 'Button: Light',
        element: 'a',
        attributes: {             
            'style': 'display: inline-block; font-weight: 400; line-height: 1.5; color: #000; background-color: #EEE; padding: 10px 20px; text-decoration: none; border-radius: 0.375rem; border: 1px solid #DDD;'
        }
    },
	{
        name: 'Button: Red',
        element: 'a',
        attributes: {             
            'style': 'display: inline-block; font-weight: 400; line-height: 1.5; color: #FFF; background-color: #DC3545; padding: 10px 20px; text-decoration: none; border-radius: 0.375rem; border: 1px solid #DC3545;'
        }
    },
    {
        name: 'Button: Outline Dark',
        element: 'a',
        attributes: {            
            'style': 'display: inline-block; font-weight: 400; line-height: 1.5; color: #212529; background-color: transparent; padding: 10px 20px; text-decoration: none; border-radius: 0.375rem; border: 1px solid #212529;'
        }
    },  

    /* --- SECTION: UTILITIES (IMAGE & BLOCK DECOR) --- */
    {
        name: 'Soft Shadow',
        element: 'div',
        attributes: { 
            'style': 'box-shadow: 0 0.5rem 1rem rgba(0, 0, 0, 0.15); border: 1px solid #eee; padding: 1rem; border-radius: 8px;' 
        }
    }
]);

CKEDITOR.config.toolbar_siteBuilder = 
[
	['Save','Preview','-','Templates','Iframe'],
	['Cut','Copy','Paste','-','Print'],
	['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat','-','Source','Maximize','ShowBlocks','-','About'],
	'/',
	['Bold','Italic','Underline','Strike','-','Subscript','Superscript'],
	['NumberedList','BulletedList','-','Outdent','Indent','Blockquote','CreateDiv'],
	['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
	['Link','Unlink','Anchor'],
	['Image','Flash','Table','HorizontalRule','Smiley','SpecialChar'],
	'/',
	['Format','Font','FontSize','Styles'],
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