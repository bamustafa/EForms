var $ = shortpoint.base.libs.$; // jQuery
var _ = shortpoint.base.libs._; // Underscore
var moment = shortpoint.base.libs.moment; // Moment.js

var _layout = '/_layouts/15/PCW/General/EForms', 
    _ImageUrl = _spPageContextInfo.webAbsoluteUrl + '/Style%20Library/tooltip.png',
    _ListInternalName = _spPageContextInfo.serverRequestPath.split('/')[5],
    _ProjectNumber = _spPageContextInfo.serverRequestPath.split('/')[2],
    _ListFullUrl = _spPageContextInfo.webAbsoluteUrl + '/Lists/' + _ListInternalName;
    _webUrl = _spPageContextInfo.webAbsoluteUrl;

let Inputelems = document.querySelectorAll('input[type="text"]');

var _modulename = "", _formType = "";
let fontSize = '17px';

let _ProjectId = '', _ProjectName = '', _WorkType = '';

const itemsToRemove = ['Status', 'State', 'Code', 'WorkflowStatus'];
const blueColor = '#6ca9d5', greenColor = '#5FC9B3', yellowColor = '#F7D46D', redColor = '#F28B82';

var appBarItems;

var onRender = async function (moduleName){ 

	try { 

        const startTime = performance.now();	

		if(moduleName == 'SPointLoad')
		{
            //debugger; 
            let isSiteADM = _spPageContextInfo.;
            const url = _spPageContextInfo.webAbsoluteUrl;
            await loadScriptAsync(`${url}${_layout}/common libs/jquery/jquery.min.js`); 
            await PreloaderScripts();
            await loadScripts();
            
            _ProjectId = localStorage.getItem(`${_ProjectNumber}-ProjectId`);
            _ProjectName = localStorage.getItem(`${_ProjectNumber}-ProjectName`);
            _WorkType = localStorage.getItem(`${_ProjectNumber}-WorkType`);
           
            if ((!_ProjectId && !_ProjectName)) {
                await fetchProjectInfoMethod(_ProjectNumber);
            }
            
            console.log(`ProjectId = ${_ProjectId}`);
            console.log(`ProjectName = ${_ProjectName}`);
            console.log(`WorkType = ${_WorkType}`);

            let LinkTitle = _ProjectId ? `${_ProjectNumber} - ${_ProjectName}` : _ProjectNumber;
            
            const iconPath = url + '/_layouts/15/Images/animdarlogo1.png';
            const linkElement = `<a href="${url}" style="text-decoration: none; color: inherit; display: flex; align-items: center;font-size: 18px;">
                                    <img src="${iconPath}" alt="Icon" style="width: 50px; height: 26px; margin-right: 14px;">                                    
                                    ${LinkTitle}
                                </a>`;

            $('span.o365cs-nav-brandingText').html(linkElement);

            // Common CSS adjustments
            const applyCommonCSS = () => {
                $('div.ms-compositeHeader-rightControls').hide(); // OK
                $('.ms-compositeHeader.root-73.shortpoint-proxy-theme-site-header').css('height', '46px'); // NO
                $('.ms-compositeHeader-mainLayout').hide(); // admin                
                //$('.Files-rightPaneInteractionContainer').css('marginTop', '-14px'); // No Action
                $('.pageTitle_dececfca').hide(); // OK
                $('.ControlZone-control').css('marginTop', '-25px');
                $('.ms-HorizontalNavItems').css('marginTop', '3px');
            };
            applyCommonCSS();

            // Prepend SVG icons to navigation items
            const icons = {
                "Home": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-home" style="margin-right: 2px;">
                <path d="M3 9L12 2l9 7v11a2 2 0 0 1-2 2h-4a2 2 0 0 1-2-2v-4H9v4a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V9z"></path>
                </svg>`,
                "Libraries": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-home" style="margin-right: 2px;transform: translateY(2px);">
                <path d="M4 4h5l2 3h9a2 2 0 0 1 2 2v7a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2z"></path>
                </svg>`,
                "Sync (Dar to Subs)": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-home" style="margin-right: 2px;">
                <path d="M23 4v6h-6"></path>
                <path d="M1 20v-6h6"></path>
                <path d="M3.51 9a9 9 0 0 1 14.36-3.36L23 10M1 14l5.63 5.63A9 9 0 0 0 20.49 15"></path>
                </svg>`,
                "MOM": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-home" style="margin-right: 2px;transform: translateY(-1px);">
                <path d="M3 3h18v4H3V3zm0 8h18v4H3v-4zm0 8h18v4H3v-4z"></path>
                </svg>`,
                "Re-Structured Files": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-home" style="margin-right: 2px;">
                <path d="M3 7V4a1 1 0 0 1 1-1h4l2 2h8a1 1 0 0 1 1 1v4H3z"></path>
                <path d="M3 11h18v8a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-8z"></path>
                </svg>`,
                "HSSE": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-home" style="margin-right: 2px;">
                <path d="M12 2L2 7v7c0 5 8 9 10 9s10-4 10-9V7l-10-5z"></path>
                <path d="M12 7v5"></path>
                <path d="M7 12h10"></path>
                </svg>`,
                "BIM Documents": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-home" style="margin-right: 2px;">
                <path d="M3 7v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V7M3 7h6l2-2h8v4M3 7h18" />
                </svg>`,
                "Quality Docs": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-folder" style="margin-right: 2px;">
                    <path d="M22 14v-2a2 2 0 0 0-2-2H12l-2-2H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-2"></path>
                </svg>`,
                "Project References and Procedures": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" style="margin-right: 2px;">
                    <path d="M4.5 0a.5.5 0 0 1 .5.5V3h5V1a.5.5 0 0 1 .5-.5h4a.5.5 0 0 1 .5.5v13a.5.5 0 0 1-.5.5h-12a.5.5 0 0 1-.5-.5V1A.5.5 0 0 1 4.5 0zm7 1h3v12h-3V1zM6 4h4v1H6V4zM6 6h4v1H6V6zM6 8h3v1H6V8zM6 10h3v1H6v-1z"/>
                    <path d="M9.5 10.5a.5.5 0 0 1-.5.5H8v1h1a.5.5 0 0 1 .5.5v1h1v-1.5a1.5 1.5 0 0 0-1.5-1.5z"/>
                </svg>`,
                "Communication": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" style="margin-right: 2px;">
                    <path d="M2 2a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H4l-2 2v-2H2a1 1 0 0 1-1-1V2zm1 0v9h10V2H3z"/>
                    <path d="M5 6h6v1H5V6zm0 2h4v1H5V8zm0 2h6v1H5v-1z"/>
                </svg>`,
                "Project Management": `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" style="margin-right: 2px;">
                    <path d="M3 2a1 1 0 0 1 1-1h8a1 1 0 0 1 1 1v1H3V2z"/>
                    <path d="M2 3h12a1 1 0 0 1 1 1v9a1 1 0 0 1-1 1H2a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1zm1 1v9h10V4H3z"/>
                    <path d="M4 6h8v1H4V6zm0 2h6v1H4V8zm0 2h8v1H4v-1z"/>
                </svg>`
            };
            
            Object.keys(icons).forEach(title => {
                $(`.ms-HorizontalNavItem-link[title="${title}"]`).prepend(icons[title]);
            });

            $('.ms-HorizontalNavItem-splitbutton').prepend(`
                <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="margin-left: 3px;">
                <path d="M6 9l6 6 6-6"></path>
                </svg>
            `);         

            // Add Background on tab mouse hoover
            const css = `
                .highlight {
                    background-color: #d6d6d6; /* Light gray background on hover */
                    color: #333;               /* Change text color */                
                }
            `;
            // Create a style element and append to head
            $('<style>').prop('type', 'text/css').html(css).appendTo('head');

            // On hover, add a highlight class
            $('.ms-HorizontalNavItem a').hover(
                function() {
                    // Add highlight class when mouse enters
                    $(this).parent('.ms-HorizontalNavItem').addClass('highlight');
                }, 
                function() {
                    // Remove highlight class when mouse leaves
                    $(this).parent('.ms-HorizontalNavItem').removeClass('highlight');
                }
            );

            $('.ms-HorizontalNavItem button').hover(
                function() {
                    // Add highlight class when mouse enters
                    $(this).parent('.ms-HorizontalNavItem').addClass('highlight');
                }, 
                function() {
                    // Remove highlight class when mouse leaves
                    $(this).parent('.ms-HorizontalNavItem').removeClass('highlight');
                }
            );
            
            // Remove the default menu arrow collapsed
            $('i[data-icon-name="ChevronDown"].ms-HorizontalNav-chevronDown').css('display', 'none');

            //Shorpoint Package icon Margin
            $('.shortpoint-icon-normal').css('marginTop', '-5px');             

            checkIfUserIsSiteAdmin();

            const elements = document.querySelectorAll('.ms-compositeHeader[class*="root-"]');

            // Set the height to 46px for each matched element
            elements.forEach(element => {
                element.style.height = '46px';
            });

            if ((_ProjectId && _ProjectName)) {          

                $('.banner_972f0ed1').css('marginLeft', '50px');
                $('.canvasWrapper_c1232e3b.shortpoint-proxy-theme-scroll-content').css('marginLeft', '50px'); 

                const customAppBarItems = {
                    'D24117-General': [
                      {
                        svgPath: `<path d="M12,14a3,3,0,0,0-3,3v7.026h6V17A3,3,0,0,0,12,14Z"/><path d="M13.338.833a2,2,0,0,0-2.676,0L0,10.429v10.4a3.2,3.2,0,0,0,3.2,3.2H7V17a5,5,0,0,1,10,0v7.026h3.8a3.2,3.2,0,0,0,3.2-3.2v-10.4Z"/>`,
                        redirectUrl: '/projects/AMI-02/SitePages/Home.aspx',
                        viewBox: '0 0 24 24',
                        iconTitle: 'GN',
                        tooltip: 'General',
                        editors: 'all',
                        readers: 'all'
                      }
                    ],
                    'D24117-0100D': [
                      {
                        svgPath: `<path d="M12.016,5.731L3.025,2.724,10.427,.257c1.026-.342,2.136-.342,3.162,0l7.419,2.473-8.992,3.001Zm-3.182,5.623l2.175-3.624L3.063,5.081c-.617-.206-1.293,.045-1.628,.602L.308,7.563c-.667,1.112-.133,2.555,1.097,2.965l5.081,1.694c.889,.296,1.865-.065,2.347-.868ZM3.008,2.718v.011l.017-.006-.017-.006ZM11.008,23.905V11.617l-.46,.766c-.742,1.236-2.044,1.945-3.415,1.945-.425,0-.856-.067-1.28-.209l-3.845-1.281v4.558c0,2.152,1.377,4.063,3.419,4.743l4.435,1.478c.374,.121,.758,.22,1.146,.287ZM23.728,7.596l-1.148-1.913c-.334-.557-1.011-.808-1.628-.602l-7.945,2.648,2.175,3.624c.482,.804,1.458,1.165,2.347,.868l5.118-1.706c1.211-.404,1.737-1.825,1.08-2.92Zm-6.845,6.733c-1.371,0-2.673-.708-3.415-1.945l-.46-.766v12.282c.422-.074,.84-.182,1.236-.314h.01l4.335-1.446c2.042-.681,3.419-2.591,3.419-4.743v-4.559l-3.845,1.282c-.424,.142-.855,.209-1.28,.209Z"/>`,
                        redirectUrl: '/projects/D24117-0100D',
                        viewBox: '0 0 24 24',
                        iconTitle: 'PK1',
                        tooltip: 'D24117-0100D',
                        editors: 'all',
                        readers: 'all'
                      }
                    ],
                    'D24117-0200D': [
                      {
                        svgPath: `<path d="M12.016,5.731L3.025,2.724,10.427,.257c1.026-.342,2.136-.342,3.162,0l7.419,2.473-8.992,3.001Zm-3.182,5.623l2.175-3.624L3.063,5.081c-.617-.206-1.293,.045-1.628,.602L.308,7.563c-.667,1.112-.133,2.555,1.097,2.965l5.081,1.694c.889,.296,1.865-.065,2.347-.868ZM3.008,2.718v.011l.017-.006-.017-.006ZM11.008,23.905V11.617l-.46,.766c-.742,1.236-2.044,1.945-3.415,1.945-.425,0-.856-.067-1.28-.209l-3.845-1.281v4.558c0,2.152,1.377,4.063,3.419,4.743l4.435,1.478c.374,.121,.758,.22,1.146,.287ZM23.728,7.596l-1.148-1.913c-.334-.557-1.011-.808-1.628-.602l-7.945,2.648,2.175,3.624c.482,.804,1.458,1.165,2.347,.868l5.118-1.706c1.211-.404,1.737-1.825,1.08-2.92Zm-6.845,6.733c-1.371,0-2.673-.708-3.415-1.945l-.46-.766v12.282c.422-.074,.84-.182,1.236-.314h.01l4.335-1.446c2.042-.681,3.419-2.591,3.419-4.743v-4.559l-3.845,1.282c-.424,.142-.855,.209-1.28,.209Z"/>`,
                        redirectUrl: '/projects/D24117-0200D',
                        viewBox: '0 0 24 24',
                        iconTitle: 'PK2',
                        tooltip: 'D24117-0200D',
                        editors: 'all',
                        readers: 'all'
                      }
                    ],
                    'D24117-0300D': [
                      {
                        svgPath: `<path d="M12.016,5.731L3.025,2.724,10.427,.257c1.026-.342,2.136-.342,3.162,0l7.419,2.473-8.992,3.001Zm-3.182,5.623l2.175-3.624L3.063,5.081c-.617-.206-1.293,.045-1.628,.602L.308,7.563c-.667,1.112-.133,2.555,1.097,2.965l5.081,1.694c.889,.296,1.865-.065,2.347-.868ZM3.008,2.718v.011l.017-.006-.017-.006ZM11.008,23.905V11.617l-.46,.766c-.742,1.236-2.044,1.945-3.415,1.945-.425,0-.856-.067-1.28-.209l-3.845-1.281v4.558c0,2.152,1.377,4.063,3.419,4.743l4.435,1.478c.374,.121,.758,.22,1.146,.287ZM23.728,7.596l-1.148-1.913c-.334-.557-1.011-.808-1.628-.602l-7.945,2.648,2.175,3.624c.482,.804,1.458,1.165,2.347,.868l5.118-1.706c1.211-.404,1.737-1.825,1.08-2.92Zm-6.845,6.733c-1.371,0-2.673-.708-3.415-1.945l-.46-.766v12.282c.422-.074,.84-.182,1.236-.314h.01l4.335-1.446c2.042-.681,3.419-2.591,3.419-4.743v-4.559l-3.845,1.282c-.424,.142-.855,.209-1.28,.209Z"/>`,
                        redirectUrl: '/projects/D24117-0300D',
                        viewBox: '0 0 24 24',
                        iconTitle: 'PK3',
                        tooltip: 'D24117-0300D',
                        editors: 'all',
                        readers: 'all'
                      }
                    ],
                    'D24117-Online': [
                      {
                        svgPath: `<path d="m13 22h-7.317a5.844 5.844 0 0 1 -5.626-4.7 5.446 5.446 0 0 1 2.885-5.651 7.652 7.652 0 0 1 -.8-5.18 8 8 0 0 1 15.49-.841 1.085 1.085 0 0 0 .721.732 8.024 8.024 0 0 1 2.98 1.674c-.11-.008-.218-.034-.333-.034a5.009 5.009 0 0 0 -4.92 4.105l-.847.424a4.953 4.953 0 0 0 -2.233-.529 5 5 0 0 0 0 10zm8-6a3 3 0 1 0 -3-3 2.9 2.9 0 0 0 .037.363l-2.96 1.481a3 3 0 1 0 0 4.312l2.96 1.481a2.9 2.9 0 0 0 -.037.363 3.015 3.015 0 1 0 .923-2.156l-2.96-1.481a2.9 2.9 0 0 0 .037-.363 2.9 2.9 0 0 0 -.037-.363l2.96-1.481a2.986 2.986 0 0 0 2.077.844z"/>`,
                        redirectUrl: 'https://darcairo.sharepoint.com/sites/D24002-0100D',
                        viewBox: '0 0 24 24',
                        iconTitle: 'Online',
                        tooltip: 'Sharepoint Online',
                        editors: 'all',
                        readers: 'all'
                      }
                    ]
                };

                const commonAppBarItems = [
                    {
                        svgPath: `
                        <!-- Top Triangle -->
                        <path d="M16.5,2.5L4,8.25L15,14L19.5,10.75L16.5,2.5Z" fill="#002050"/>
                        
                        <!-- Bottom Triangle -->
                        <path d="M4,8.25L4,21.5L15,14L4,8.25Z" fill="#002050"/>
            
                        <!-- Inner Triangle -->
                        <path d="M7,11L4,21.5L15,14L7,11Z" fill="#ffffff"/>
                      `,
                        redirectUrl: 'https://ax.d365.dar.com/namespaces/AXSF/?mi=ProjProjectsListPage',
                        viewBox: '0 0 24 24',
                        iconTitle: 'D365',
                        tooltip: 'Manage project on D365',
                        editors: 'all',
                        readers: 'all'
                    }
                ];           

                let projectNumbers = [];

                if(url.toLowerCase().includes('d24117') || url.toLowerCase().includes('ami-02')){
                    projectNumbers = ['D24117-General', 'D24117-0100D', 'D24117-0200D', 'D24117-0300D', 'D24117-Online'];
                }
                appBarItems = getAppBarItemsForProjects(projectNumbers, customAppBarItems, commonAppBarItems);

                appBar();               
            
                // #region Design Projects MENU BAR 
                if(_WorkType.toLowerCase() === 'design'){
                
                    // Project References and Procedures Tab
                    addHorizontalNavItemWithDropdown(
                        'Project References and Procedures', 
                        '#', 
                        `<path d="M11 2.5a2.5 2.5 0 1 1 .603 1.628l-6.718 3.12a2.5 2.5 0 0 1 0 1.504l6.718 3.12a2.5 2.5 0 1 1-.488.876l-6.718-3.12a2.5 2.5 0 1 1 0-3.256l6.718-3.12A2.5 2.5 0 0 1 11 2.5"/>`, 
                        [
                            { title: 'Applications', href: `${_webUrl}/PRPs/Applications` },
                            { title: 'Collected Data', href: `${_webUrl}/PRPs/Collected%20Data` },
                            { title: 'Issued Tender Documents & Notices', href: `${_webUrl}/PRPs/Issued%20Tender%20Documents%20and%20Notices` },
                            { title: 'Manuals & Procedures', href: `${_webUrl}/PRPs/Manuals%20and%20Procedures` },
                            { title: 'Other', href: `${_webUrl}/PRPs/Other` },
                            { title: 'Proposal Documents', href: `${_webUrl}/PRPs/Proposal%20Documents` },
                            { title: 'Terms of Reference', href: `${_webUrl}/PRPs/Terms%20of%20Reference` }
                        ]
                    );

                    // Communication Tab
                    addHorizontalNavItemWithDropdown(
                        'Communication', 
                        '#', 
                        `<path d="M0 4a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2zm9 1.5a.5.5 0 0 0 .5.5h4a.5.5 0 0 0 0-1h-4a.5.5 0 0 0-.5.5M9 8a.5.5 0 0 0 .5.5h4a.5.5 0 0 0 0-1h-4A.5.5 0 0 0 9 8m1 2.5a.5.5 0 0 0 .5.5h3a.5.5 0 0 0 0-1h-3a.5.5 0 0 0-.5.5m-1 2C9 10.567 7.21 9 5 9c-2.086 0-3.8 1.398-3.984 3.181A1 1 0 0 0 2 13h6.96q.04-.245.04-.5M7 6a2 2 0 1 0-4 0 2 2 0 0 0 4 0"/>`, 
                        [
                            //{ title: 'Calendar', href: `${_webUrl}/Lists/Calendar/Calendar.aspx` },
                            { title: 'Central Files', href: `${_webUrl}/SitePages/CentralFiling.aspx` },
                            //{ title: 'Contacts', href: `${_webUrl}/Lists/Contacts/AllItems.aspx` },
                            { title: 'Correspondences', href: `${_webUrl}/Correspondences/Forms/AllItems.aspx` },
                            { title: 'Minutes of Meeting', href: `${_webUrl}/Minutes of Meeting/Forms/AllItems.aspx` },                     
                            //{ title: 'Notices', href: `${_webUrl}/Lists/Notices/AllItems.aspx` } //,
                            //{ title: 'Workflow Correspondences List', href: '#subitem6' }
                        ]
                    );

                    // Project Management Tab
                    addHorizontalNavItemWithDropdown(
                        'Project Management', 
                        '#', 
                        `<path d="M3 2a1 1 0 0 1 1-1h8a1 1 0 0 1 1 1v1H3V2z"/>
                        <path d="M2 3h12a1 1 0 0 1 1 1v9a1 1 0 0 1-1 1H2a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1zm1 1v9h10V4H3z"/>
                        <path d="M4 6h8v1H4V6zm0 2h6v1H4V8zm0 2h8v1H4v-1z"/>`, 
                        [                        
                            { title: 'Confidential', href: `${_webUrl}/Confidential/Forms/AllItems.aspx` },
                            { 
                                title: 'Project Management Folders', 
                                href: '#subitem4',
                                children: [
                                    { title: 'Checklist for Project Submittal', href: `${_webUrl}/Project Management/Checklist%20for%20Project%20Submittal` },
                                    { title: 'Customer Feedback', href: `${_webUrl}/Project Management/Customer%20Feedback` },
                                    { title: 'Leassons Learned Register', href: `${_webUrl}/Project Management/Lessons%20Learned%20Register` },
                                    { title: 'Monthly Progress Reports', href: `${_webUrl}/Project Management/Monthly%20Progress%20Reports` },
                                    { title: 'Organization Chart', href: `${_webUrl}/Project Management/Organization%20Chart` },
                                    { title: 'Schedule', href: `${_webUrl}/Project Management/Schedule` },
                                    { title: 'Supplier Control', href: `${_webUrl}/Project Management/Supplier%20Control` } 
                            ]},                                          
                            { title: 'DRD Guidelines', href: `${_webUrl}/SitePages/PlumsailForms/DRDGuidelines/Item/NewForm.aspx` },
                            { 
                                title: 'RACI', 
                                href: '#subitem1', 
                                children: [
                                { title: 'Change Log', href: `${_webUrl}/Lists/Change Log/AllItems.aspx` },
                                { title: 'Design Codes & Regulations', href: `${_webUrl}/Lists/LocalDesignCode/AllItems.aspx` },
                                { title: 'Risk Register', href: `${_webUrl}/Lists/RiskRegister/AllItems.aspx` }
                            ]},                     
                            { title: 'Restricted', href: `${_webUrl}/Restricted/Forms/AllItems.aspx` }                    
                            //{ title: 'Workflow Correspondences List', href: '#subitem6' }
                        ]
                    );

                    // Design Workspace Tab
                    addHorizontalNavItemWithDropdown(
                        'Design Workspace', 
                        '#', 
                        `<path d="M4 16s-1 0-1-1 1-4 5-4 5 3 5 4-1 1-1 1zm4-5.95a2.5 2.5 0 1 0 0-5 2.5 2.5 0 0 0 0 5"/>
                        <path d="M2 1a2 2 0 0 0-2 2v9.5A1.5 1.5 0 0 0 1.5 14h.653a5.4 5.4 0 0 1 1.066-2H1V3a1 1 0 0 1 1-1h12a1 1 0 0 1 1 1v9h-2.219c.554.654.89 1.373 1.066 2h.653a1.5 1.5 0 0 0 1.5-1.5V3a2 2 0 0 0-2-2z"/>`, 
                        [
                            { title: 'Shared', href: `${_webUrl}/Shared/Forms/AllItems.aspx` },
                            { title: 'Sustainability Documents', href: `${_webUrl}/Sustainability Documents/Forms/AllItems.aspx` },
                            { title: 'Work In Progress', href: `${_webUrl}/WIP/Forms/AllItems.aspx` }                        
                        ]
                    );

                    // Departmental Project Filing Tab
                    addHorizontalNavItemWithDropdown(
                        'Departmental Project Filing', 
                        '#', 
                        `<path d="M12 0H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V2a2 2 0 0 0-2-2m-1.146 6.854-3 3a.5.5 0 0 1-.708 0l-1.5-1.5a.5.5 0 1 1 .708-.708L7.5 8.793l2.646-2.647a.5.5 0 0 1 .708.708"/>`, 
                        [
                            { title: 'BIM Internal Checking', href: `${_webUrl}/Design%20Project%20Filing/BIM%20Internal%20Checking` },
                            { title: 'Departmental Tender Documents', href: `${_webUrl}/Design%20Project%20Filing/Departmental%20Tender%20Documents` },
                            { title: 'Design and Safety Review', href: `${_webUrl}/Design%20Project%20Filing/Design%20and%20Safety%20Review` },
                            { title: 'Design Calculation', href: `${_webUrl}/Design%20Project%20Filing/Design%20Calculation` },
                            { title: 'Engineering Reports and Studies', href: `${_webUrl}/Design%20Project%20Filing/Engineering%20Reports%20and%20Studies` },  
                            { title: 'Interdepartmental Checking', href: `${_webUrl}/Design%20Project%20Filing/Interdepartmental%20Checking` }                     
                        ]
                    );

                    // Departmental Project Filing Tab
                    addHorizontalNavItemWithDropdown(
                        'Quality Management', 
                        '#', 
                        `<path d="M10.854 8.146a.5.5 0 0 1 0 .708l-3 3a.5.5 0 0 1-.708 0l-1.5-1.5a.5.5 0 0 1 .708-.708L7.5 10.793l2.646-2.647a.5.5 0 0 1 .708 0"/>
                        <path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5M2 2a1 1 0 0 0-1 1v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V3a1 1 0 0 0-1-1z"/>
                        <path d="M2.5 4a.5.5 0 0 1 .5-.5h10a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5H3a.5.5 0 0 1-.5-.5z"/>`, 
                        [
                            { title: 'Audit Reports', href: `${_webUrl}/Quality%20Management/Audit%20Reports%20-%20CARs` },
                            { title: 'Audit Schedule', href: `${_webUrl}/Quality%20Management/Audit%20Schedule` }                                            
                        ]
                    );

                    // Design Deliverables Tab
                    addHorizontalNavItemWithDropdown(
                        'Design Deliverables', 
                        '#', 
                        `<path d="M9.828 3h3.982a2 2 0 0 1 1.992 2.181l-.637 7A2 2 0 0 1 13.174 14H2.825a2 2 0 0 1-1.991-1.819l-.637-7a2 2 0 0 1 .342-1.31L.5 3a2 2 0 0 1 2-2h3.672a2 2 0 0 1 1.414.586l.828.828A2 2 0 0 0 9.828 3m-8.322.12q.322-.119.684-.12h5.396l-.707-.707A1 1 0 0 0 6.172 2H2.5a1 1 0 0 0-1 .981z"/>`, 
                        [
                            { title: 'Initiate Submittals', href: `${_webUrl}/Lists/InitiateSubmittals/AllItems.aspx` },
                            { title: 'Design Tasks', href: `${_webUrl}/Lists/DesignTasks/ActiveTasks.aspx` },
                            { title: 'TentativeLOD', href: `${_webUrl}/TentativeLOD/Forms/AllItems.aspx` },
                            { title: 'Deliverables', href: `${_webUrl}/Deliverables/Forms/AllItems.aspx` },
                            { title: 'RLOD', href: `${_webUrl}/Lists/RLOD/AllItems.aspx` },
                            { title: 'SLOD', href: `${_webUrl}/Lists/SLOD/AllItems1.aspx` }                                       
                        ]
                    );

                    // Tender Management Tab
                    addHorizontalNavItemWithDropdown(
                        'Tender Management', 
                        '#', 
                        `<path d="M9.405 1.05c-.413-1.4-2.397-1.4-2.81 0l-.1.34a1.464 1.464 0 0 1-2.105.872l-.31-.17c-1.283-.698-2.686.705-1.987 1.987l.169.311c.446.82.023 1.841-.872 2.105l-.34.1c-1.4.413-1.4 2.397 0 2.81l.34.1a1.464 1.464 0 0 1 .872 2.105l-.17.31c-.698 1.283.705 2.686 1.987 1.987l.311-.169a1.464 1.464 0 0 1 2.105.872l.1.34c.413 1.4 2.397 1.4 2.81 0l.1-.34a1.464 1.464 0 0 1 2.105-.872l.31.17c1.283.698 2.686-.705 1.987-1.987l-.169-.311a1.464 1.464 0 0 1 .872-2.105l.34-.1c1.4-.413 1.4-2.397 0-2.81l-.34-.1a1.464 1.464 0 0 1-.872-2.105l.17-.31c.698-1.283-.705-2.686-1.987-1.987l-.311.169a1.464 1.464 0 0 1-2.105-.872zM8 10.93a2.929 2.929 0 1 1 0-5.86 2.929 2.929 0 0 1 0 5.858z"/>`, 
                        [
                            { title: 'Bidders', href: `${_webUrl}/Lists/Bidders/AllItems.aspx` },
                            { title: 'TenderLibrary', href: `${_webUrl}/TenderLibrary/Forms/AllItems.aspx` },
                            { title: 'QApp', href: `${_webUrl}/Lists/Tender Query/AllItems.aspx` }                                       
                        ]
                    );
                }
                //#endregion
            }
        }      

        const endTime = performance.now();
        const elapsedTime = endTime - startTime;
        console.log(`Execution time SPointLoad: ${elapsedTime} milliseconds`);
	}
	catch (e) {		
		console.log(e);	
	}

    preloader("remove");
}

//#region Additional Function

async function loadingButtons(){  

    fd.toolbar.buttons.push({
        icon: 'CheckMark',
        class: 'btn-outline-primary',
        disabled: false,
        text: 'Acknowledge',
        style: `background-color:${greenColor}; color:white; width:200px !important`,
        click: async function() {  	
            if(fd.isValid){
                await PreloaderScripts();          
                fd.field('Acknowledged').value = true;
                fd.field('LastAccessedDate').value = new Date();
                fd.save();
            }            
        }
    });   
    
    fd.toolbar.buttons.push({
        icon: 'Cancel',
        class: 'btn-outline-primary',
        text: 'Cancel',	
        style: `background-color:${redColor}; color:white; width:200px !important`,
        click: async function() {
            await PreloaderScripts();
            fd.close();
        }			
	});           
}

async function processClosedWindowResult() {

    const folderUrl = `${_spPageContextInfo.siteServerRelativeUrl}/CORLibrary/${_refNo}`;    
    try {        
        const files = await pnp.sp.web.getFolderByServerRelativeUrl(folderUrl).files();
        return files.length;
    } catch (error) {
        console.error("Error fetching files:", error);
    }
}

function formatingButtonsBar(){
    
    $('div.ms-compositeHeader').remove()
    $('span.o365cs-nav-brandingText').text(`${_ProjectNumber} - Job Architecture Form`);
    $('i.ms-Icon--PDF').remove();
          
    let toolbarElements = document.querySelectorAll('.fd-toolbar-primary-commands');
    toolbarElements.forEach(function(toolbar) {
        toolbar.style.display = "flex";
        //toolbar.style.justifyContent = "flex-end";
        toolbar.style.marginLeft = "44px";            
    });

    let commandBarElement = document.querySelectorAll('[aria-label="Command Bar."]');
        commandBarElement.forEach(function(element) {        
        element.style.paddingTop = "16px";       
    }) ;
    
    document.querySelectorAll('.CanvasZoneContainer.CanvasZoneContainer--read').forEach(element => {
        element.style.marginTop = '7px';
        element.style.marginLeft = '-100px';
    });

    var fieldTitleElements = document.querySelectorAll('.fd-form .row > .fd-field-title');

    fieldTitleElements.forEach(function(element) {
        element.style.fontWeight = 'bold';      
        element.style.borderTopLeftRadius = '6px';
        element.style.borderBottomLeftRadius = '6px';
        element.style.width = '200px';
        element.style.display = 'inline-block';
    });

    document.querySelectorAll(".homepageContainer_ce0ab33e.scrollRegion_ce0ab33e.homepageContainerNewDigestNavBar_ce0ab33e").forEach(element => {
        element.style.textAlign = "justify";
    });

    document.querySelectorAll('.homepageContainer_ce0ab33e.scrollRegion_ce0ab33e.homepageContainerNewDigestNavBar_ce0ab33e').forEach(element => {
        element.style.backgroundColor = "#f4f4f4";
    });

    document.querySelectorAll('.ControlZone').forEach(element => {
        // Uncomment the line below if you want to use the gradient background
        // element.style.background = "linear-gradient(to right, rgb(218, 237, 216), rgb(187, 229, 218), rgb(158, 214, 224), rgb(150, 182, 235), rgb(175, 169, 240), rgb(175, 168, 240), rgb(168, 165, 239))";
                
        element.style.background = "#ffffff";
    });
}

function ReadOnly(elem){ 
    $(elem).prop("readonly", true).css({
	    "background-color": "transparent",
	    "border": "0px"
	});  
}

function fixTextArea(){
	$("textarea").each(function(index){		
		$(this).css('height', '150px');
		var height = (this.scrollHeight + 5) + "px";
        $(this).css('height', height);
	});
}

function FixListTabelRows(){ 
    
    let tables = $("table[role='grid']");
    tables.each(function(tblIndex, tbl){
        $(tbl).find('tr').each(function(trIndex, tr) {
    	  
    	    if (trIndex === 0 ){    	
    		   let childs = tr.children;
    		   if(childs.length > 0){
    		     childs[0].style.textAlign = 'center';
    			 childs[1].style.textAlign = 'center';
                }                  		   
    		}
    		
    	   $(tr).find('td').each(function(tdIndex, td) {
                let $td = $(td);
                
                if (tdIndex === 0 || tdIndex === 1)
                    td.style.textAlign = 'center';
                    
                else{
                    if(_formType !== 'Display')
                        $td.children().css('whiteSpace', 'nowrap');
                }
                 
                if(_formType !== 'Display')
                    $td.css('whiteSpace', 'nowrap');
    		});                			
        });
    });    
}

function FixWidget(dt){
    FixListTabelRows();
    var Clientwidth = dt.$el.clientWidth; 
    //Clientwidth = Clientwidth * 96 / 100;  
    var Rwidget = dt.widget;
    var columns = Rwidget.columns;  
    var ColumnsLength = columns.length;
    var width = Clientwidth/(ColumnsLength-1);
    
    var RemainingWidth = 0;
    var RemainingWidth2 = 0;
    var RemainingWidth3 = 0;
    var RemainingWidth4 = 0;
    var RemainingWidth5 = 0;
    var RemainingWidth6 = 0;
            
    for (let i = 1; i < ColumnsLength; i++) {
    
        var field = columns[i].field;	
        
        if(field === 'Reviewed'){
            var ReviewedWidth = 80;
            RemainingWidth = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        }  
        else if(field === 'LinkTitle'){
            var ReviewedWidth = 350;
            RemainingWidth2 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        } 
        else if(field === 'FieldOfSpecialization'){
            var ReviewedWidth = 180;
            RemainingWidth3 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        } 
        else if(field === 'Qualifications'){
            var ReviewedWidth = 400;
            RemainingWidth4 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        }  
        else if(field === 'From'){
            var ReviewedWidth = 90;
            RemainingWidth5 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        } 
        else if(field === 'To'){
            var ReviewedWidth = 90;
            RemainingWidth6 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        } 
        else if(i == (ColumnsLength - 1)){                  
            dt._columnWidthViewStorage.set(field, (width + RemainingWidth + RemainingWidth2 + RemainingWidth3 + RemainingWidth4 + RemainingWidth5 + RemainingWidth6)); 
            Rwidget.resizeColumn(columns[i], (width + RemainingWidth + RemainingWidth2 + RemainingWidth3 + RemainingWidth4 + RemainingWidth5 + RemainingWidth6)); 
        }
        else{
            dt._columnWidthViewStorage.set(field, width); 
            Rwidget.resizeColumn(columns[i], width); 
        } 
    }

    const gridContent = dt.$el.querySelector('.k-grid-content.k-auto-scrollable');
    if (gridContent) {
        gridContent.style.overflowX = 'hidden';
    }
    
    var rows = Rwidget._data;
    
    // #32DAC4
    for (let i = 0; i < rows.length; i++) {
        var Reviewed = rows[i].Reviewed; 
        if(Reviewed === 'Yes'){
            const row = $(dt.$el).find('tr[data-uid="' + rows[i].uid+ '"');
            row[0].style.background = 'linear-gradient(to right, rgb(245, 255, 245), rgb(235, 250, 245), rgb(220, 255, 250), rgb(210, 220, 255), rgb(215, 205, 255), rgb(215, 205, 255), rgb(205, 205, 255))';
        }           
    }        
}

async function CheckifUserinSPGroup() {

	let IsTMUser = "User"; 

	try{
		 await pnp.sp.web.currentUser.get()
         .then(async function(user){
			await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
			 .then(async function(groupsData){
				for (var i = 0; i < groupsData.length; i++) {				
					if(groupsData[i].Title === "MC_Reviewer")
					{					
					   IsTMUser = "MC_Reviewer";
                       _DipN = user.Title;
					   break;
				    }					
				}				
			});
	     });
    }
	catch(e){
        console.log(e);
    }
	return IsTMUser;				
}

function GetHTMLBody(EmailbodyHeader, ToName, DoneBy){	

	var Body = "<html>";
	Body += "<head>";
	Body += "<meta charset='UTF-8'>";
	Body += "<meta name='viewport' content='width=device-width, initial-scale=1.0'>";				
	Body += "</head>";
	Body += "<body style='font-family: Verdana, sans-serif; font-size: 12px; line-height: 1.5; color: #333;'>";
	Body += "<div style='max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ccc; border-radius: 5px; background-color: #f9f9f9;'>"; 	
    
    Body += "<div style='margin-top: 10px;'>";
	Body += "<p style='margin: 0 0 10px;'>Dear <strong>" + ToName + "</strong>,</p>";
    Body += '</br>' ;
    Body += EmailbodyHeader ;
    Body += "<p style='margin: 0 0 10px;'>Best regards,</p>";
	Body += "<p style='margin: 0 0 10px;'><strong>" + DoneBy + "</strong></p>";
	Body += "</div>";				
	Body += "</div>";
	Body += "</body>";
	Body += "</html>";

	return Body;
}

function htmlEncode(str) {
    return str.replace(/[&<>"']/g, function(match) {
        return {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;'
        }[match];
    });
}

async function loadScripts(){
	// const libraryUrls = [		
	// 	_layout + '/controls/tooltipster/jquery.tooltipster.min.js',
	// 	_layout + '/plumsail/js/commonUtils.js'        
	// ];
  
	// const cacheBusting = '?t=' + new Date().getTime();
	//   libraryUrls.map(url => { 
	// 	  $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
	// 	});
		
	const stylesheetUrls = [
		//_layout + '/controls/tooltipster/tooltipster.css',
        _layout + '/plumsail/css/ShortPointStyle.css'      		
	];
  
	stylesheetUrls.map((item) => {
	  var stylesheet = item;
	  $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
	});

    // const fontStyles = `
	// 	@font-face {
	// 		font-family: 'SegoeUIRegular';
	// 		src: url('${_layout}/plumsail/js/FabricIcons/segoeui-regular.woff2') format('woff2'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/segoeui-regular.woff') format('woff');
	// 	}
	// 	@font-face {
	// 		font-family: 'SegoeUILight';
	// 		src: url('${_layout}/plumsail/js/FabricIcons/segoeui-light.woff2') format('woff2'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/segoeui-light.woff') format('woff');
	// 	}
	// 	@font-face {
	// 		font-family: 'FabricIcons53';
	// 		src: url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.53.woff2') format('woff2'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.53.woff') format('woff'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.53.ttf') format('truetype');
	// 	}
	// 	@font-face {
	// 		font-family: 'FabricIcons23';
	// 		src: url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.23.woff2') format('woff2'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.23.woff') format('woff'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.23.ttf') format('truetype');
	// 	}
	// 	@font-face {
	// 		font-family: 'FabricIcons354';
	// 		src: url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-3.54.woff') format('woff');
	// 	}
	// `;

	// // Append font styles to head
	// $('<style>')
	// 	.prop('type', 'text/css')
	// 	.html(fontStyles)
	// 	.appendTo('head');
}

async function PreloaderScripts(){
  
	await loadScriptAsync(_layout + '/controls/preloader/jquery.dim-background.min.js')
		.then(() => {
			return loadScriptAsync(_layout + '/plumsail/js/preloader.js');
		})
		.then(() => {
			preloader();
		});	    
}

var setButtonToolTip = async function(_btnText, toolTipMessage){  
        
    var btnElement = $('span').filter(function(){ return $(this).text() == _btnText; }).prev();
	if(btnElement.length === 0)
	  btnElement = $(`button:contains('${_btnText}')`);
	
    if(btnElement.length > 0){
	  if(btnElement.length > 1)
		btnElement = btnElement[1].parentElement;
      else btnElement = btnElement[0].parentElement;
	  
      $(btnElement).attr('title', toolTipMessage);

      $(btnElement).tooltipster({
        delay: 100,
        maxWidth: 350,
        speed: 500,
        interactive: true,
        animation: 'fade', //fade, grow, swing, slide, fall
        trigger: 'hover'
      });
    }
}

var checkIfUserIsSiteAdmin = function () {
    try {
        // const response = await fetch(`${_spPageContextInfo.webAbsoluteUrl}/_api/web/currentuser`, {
        //     method: "GET",
        //     headers: {
        //         "Accept": "application/json;odata=verbose"
        //     }
        // });

        // if (!response.ok) {
        //     throw new Error(`HTTP error! status: ${response.status}`);
        // }

        // const data = await response.json();
        // const isSiteAdmin = data.d.IsSiteAdmin; 

        var isSiteAdmin = _spPageContextInfo.isSiteAdmin;
        
        if (isSiteAdmin) {
            console.log("The current user is a site admin.");            
        } else { 
            $('.ControlZone-control').css('marginTop', '-32px');
            $('.ms-CommandBar.root_75be230e.shortpoint-proxy-neutral-lighter--bg.shortpoint-proxy-neutral-primary--text.shortpoint-proxy-theme-command-bar').css('display', 'none');
            $('.commandBarWrapper').css('display', 'none');                 
            console.log("The current user is Normal User.");
        }
    } catch (error) {
        console.error("Error checking if user is site admin: ", error.message);
    }
}

async function loadScriptAsync(src) {
    return new Promise((resolve, reject) => {
        var script = document.createElement('script');
        script.type = 'text/javascript';
        script.src = src;

        script.onload = function() {
            resolve(); // Resolve the promise when the script loads successfully
        };

        script.onerror = function() {
            console.error('Error loading script:', src); // Log an error if the script fails to load
            reject(new Error(`Failed to load script: ${src}`)); // Reject the promise on error
        };

        document.head.appendChild(script);
    });
}

function addHorizontalNavItemWithDropdown(title, href, path, subItems = []) {
    // var navItemsContainer = document.querySelector('.ms-HorizontalNavItems');
    // var newNavItem = document.createElement('span');
    // newNavItem.className = 'ms-HorizontalNavItem';

    var navItemsContainer = document.querySelector('.ms-compositeHeader-topWrapper.noNav.ms-hiddenMdDown');
    var newNavItem = document.createElement('span');
    newNavItem.className = 'ms-HorizontalNavItem';

    // Main button for the nav item
    var mainButton = `
        <button class="ms-HorizontalNavItem-link link-81 linkUnselected-82" title="${title}" aria-label="${title}" aria-haspopup="${subItems.length > 0}" role="menuitem" tabindex="0">
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" style="margin-right: 2px;">
                ${path}
            </svg>
            ${title}
        </button>`;

    // Split button for dropdown trigger (optional)
    // var splitButton = `
    //     <button class="ms-HorizontalNavItem-splitbutton splitButton-84" aria-label="this menu item with submenu. Use arrow keys and enter to navigate" aria-haspopup="true" tabindex="-1">
    //         <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="margin-left: 3px;">
    //             <path d="M6 9l6 6 6-6"></path>
    //         </svg>
    //     </button>`;

    // Add the main button and split button to the nav item
    newNavItem.innerHTML = mainButton; // + splitButton;

    // If there are sub-items, create a dropdown for them using <ul> and <li>
    if (subItems.length > 0) {
        var dropdown = document.createElement('ul');
        dropdown.className = 'dropdown-menu';
        dropdown.style.display = 'none'; // Hidden by default
        dropdown.style.position = 'absolute'; // Positioning
        dropdown.style.listStyle = 'none'; // No bullet points
        dropdown.style.padding = '0';
        dropdown.style.margin = '5px 0 0 0';
        dropdown.style.backgroundColor = '#fff'; // Background color for dropdown
        dropdown.style.border = '1px solid #ddd'; // Border
        dropdown.style.boxShadow = '0 4px 8px rgba(0, 0, 0, 0.1)'; // Light shadow        
        dropdown.style.zIndex = '9999';  // Ensure it's on top 

        // Create the icon element (use any icon or text for collapse/uncollapse)
        var dropdowncollapseIcon = document.createElement('span');
        dropdowncollapseIcon.textContent = ' ▸'; // Change icon to forward arrow
        dropdowncollapseIcon.style.cursor = 'pointer';
        dropdowncollapseIcon.style.marginLeft = '2px'; // Space between text and icon
        dropdowncollapseIcon.style.fontSize = '18px'; // Correctly apply font-size

        // Append the icon to the link
        newNavItem.appendChild(dropdowncollapseIcon);

        subItems.forEach(function(subItem) {
            var subItemElement = document.createElement('li');
            subItemElement.className = 'dropdown-item';
            subItemElement.style.position = 'relative'; // To position nested dropdown
            subItemElement.style.padding = '8px 16px'; // Padding
            subItemElement.style.cursor = 'pointer';

            // Add hover effect
            subItemElement.addEventListener('mouseover', function() {
                subItemElement.style.backgroundColor = '#f0f0f0'; // Highlight parent on hover
            });
            subItemElement.addEventListener('mouseout', function() {
                subItemElement.style.backgroundColor = ''; // Reset background color
            });

            // Add sub-item link
            var subItemLink = document.createElement('a');
            subItemLink.href = subItem.href;
            subItemLink.className = 'dropdown-link';
            subItemLink.style.textDecoration = 'none';
            subItemLink.style.color = '#333'; // Text color
            subItemLink.textContent = subItem.title;

            subItemElement.appendChild(subItemLink);

            // If the sub-item has child sub-items, create nested dropdown
            if (subItem.children && subItem.children.length > 0) {

                 // Create the icon element (use any icon or text for collapse/uncollapse)
                var collapseIcon = document.createElement('span');
                collapseIcon.textContent = ' ▾'; // Example: arrow icon
                collapseIcon.style.cursor = 'pointer';
                collapseIcon.style.marginLeft = '2px'; // Space between text and icon
                collapseIcon.style.fontSize = '16px'; // Correctly apply font-size

                // Append the icon to the link
                subItemLink.appendChild(collapseIcon);

                // Create the nested dropdown menu (initially hidden)

                var nestedDropdown = document.createElement('ul');
                nestedDropdown.className = 'nested-dropdown-menu';
                nestedDropdown.style.display = 'none'; // Hidden by default
                nestedDropdown.style.position = 'absolute';
                nestedDropdown.style.left = '100%'; // Position next to parent
                nestedDropdown.style.top = '0';
                nestedDropdown.style.listStyle = 'none';
                nestedDropdown.style.padding = '0';
                nestedDropdown.style.margin = '0';
                nestedDropdown.style.backgroundColor = '#fff';
                nestedDropdown.style.border = '1px solid #ddd';
                nestedDropdown.style.boxShadow = '0 4px 8px rgba(0, 0, 0, 0.1)';
                nestedDropdown.style.width = 'max-content'; //'max-content';                

                subItem.children.forEach(function(nestedSubItem) {
                    var nestedSubItemElement = document.createElement('li');
                    nestedSubItemElement.className = 'nested-dropdown-item';
                    nestedSubItemElement.style.padding = '8px 16px';
                    nestedSubItemElement.style.cursor = 'pointer';
                    //nestedSubItemElement.style.width = '100%'; //'max-content';

                    // Add hover effect
                    nestedSubItemElement.addEventListener('mouseover', function() {                        
                        nestedSubItemElement.style.backgroundColor = '#f0f0f0'; // Highlight on hover
                    });
                    nestedSubItemElement.addEventListener('mouseout', function() {                        
                        nestedSubItemElement.style.backgroundColor = ''; // Reset
                    });

                    var nestedSubItemLink = document.createElement('a');
                    nestedSubItemLink.href = nestedSubItem.href;
                    nestedSubItemLink.className = 'nested-dropdown-link';
                    nestedSubItemLink.style.textDecoration = 'none';
                    nestedSubItemLink.style.color = '#333';
                    nestedSubItemLink.textContent = nestedSubItem.title;

                    nestedSubItemElement.appendChild(nestedSubItemLink);
                    nestedDropdown.appendChild(nestedSubItemElement);
                });

                subItemElement.appendChild(nestedDropdown);

                // Show nested dropdown on hover
                // subItemElement.addEventListener('mouseover', function() {
                //     nestedDropdown.style.display = 'block'; // Show on hover
                // });

                // subItemElement.addEventListener('mouseout', function() {
                //     nestedDropdown.style.display = 'none'; // Hide on mouse out
                // });  
                
                // Toggle the dropdown visibility when the icon is clicked

                // Show nested dropdown on hover and change icon to collapse
                subItemElement.addEventListener('mouseover', function(e) {
                    nestedDropdown.style.display = 'block'; // Show dropdown on hover
                    collapseIcon.textContent = ' ▸'; // Change icon to forward arrow
                });

                // Hide the dropdown on mouse out and change icon to uncollapse
                subItemElement.addEventListener('mouseout', function(e) {
                    nestedDropdown.style.display = 'none'; // Hide dropdown on mouse out
                    collapseIcon.textContent = ' ▾'; // Change icon back to "uncollapse" (downward arrow)
                });
            } 
            
            // Make the entire <li> clickable
            subItemElement.addEventListener('click', function() {
                window.location.href = subItem.href; // Navigate to the href
            });

            dropdown.appendChild(subItemElement);
        });

        newNavItem.appendChild(dropdown);

        // Toggle dropdown visibility on main button click (in case on click to open the menu bar)
        // newNavItem.querySelector('.ms-HorizontalNavItem-link').addEventListener('click', function() {
        //     if (dropdown.style.display === 'none') {
        //         dropdown.style.display = 'block'; // Show the dropdown
        //     } else {
        //         dropdown.style.display = 'none'; // Hide the dropdown
        //     }
        // });

        //#region on Mouse Over Menu Top Bar
        // Show the dropdown when hovering over the navigation link
        newNavItem.querySelector('.ms-HorizontalNavItem-link').addEventListener('mouseenter', function() {
            dropdown.style.display = 'block'; // Show the dropdown on hover
            dropdowncollapseIcon.textContent = ' ▾'; // Example: arrow icon
        });

        // Keep the dropdown visible when hovering over the dropdown itself
        dropdown.addEventListener('mouseenter', function() {
            dropdown.style.display = 'block'; // Ensure dropdown stays visible when hovering over it
            dropdowncollapseIcon.textContent = ' ▾'; // Change icon back to "uncollapse" (downward arrow)
        });

        // Hide the dropdown when leaving the navigation link
        newNavItem.querySelector('.ms-HorizontalNavItem-link').addEventListener('mouseleave', function() {
            setTimeout(function() {
                // Only hide if not hovering over the dropdown
                if (!dropdown.matches(':hover')) {
                    dropdown.style.display = 'none';
                    dropdowncollapseIcon.textContent = ' ▸'; // Change icon to forward arrow
                }
            }, 200); // Small delay to prevent flickering
        });

        // Hide the dropdown when leaving the dropdown itself
        dropdown.addEventListener('mouseleave', function() {
            setTimeout(function() {
                // Only hide if not hovering over the nav item
                if (!newNavItem.querySelector('.ms-HorizontalNavItem-link').matches(':hover')) {
                    dropdown.style.display = 'none';
                    dropdowncollapseIcon.textContent = ' ▸'; // Change icon to forward arrow
                }
            }, 200); // Small delay to prevent flickering
        });
        //#endregion

        // This code for the arrow icon beside the menu tab
        // newNavItem.querySelector('.ms-HorizontalNavItem-splitbutton').addEventListener('click', function() {
        //     if (dropdown.style.display === 'none') {
        //         dropdown.style.display = 'block'; // Show the dropdown
        //     } else {
        //         dropdown.style.display = 'none'; // Hide the dropdown
        //     }
        // });

        // Close dropdown when clicking outside of it
        document.addEventListener('click', function(event) {
            if (!newNavItem.contains(event.target)) {
                dropdown.style.display = 'none'; // Hide dropdown
            }
        });        
    }

    navItemsContainer.appendChild(newNavItem);
}

var fetchProjectInfoMethod = async function (ProjectNo){
	try {              
        const restUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx`
        const serviceUrl = `${restUrl}?command=GetProjectInfo&ProjectNumber=${ProjectNo}`;                               
        let ProjectTeam = await fetchProjectTeam('GET', serviceUrl, true);             

        const GetProjectTeamNodes = ProjectTeam.getElementsByTagName("Table1");      

        for (let i = 0; i < GetProjectTeamNodes.length; i++) {

            const projectNode = GetProjectTeamNodes[i];

            // Extracting data from each projectNode
            _ProjectId = projectNode.getElementsByTagName("ProjectId")[0]?.textContent || '';
            _ProjectName = projectNode.getElementsByTagName("ProjectName")[0]?.textContent || ''; 
            _WorkType = projectNode.getElementsByTagName("WorkType")[0]?.textContent || ''; 

            localStorage.setItem(`${ProjectNo}-ProjectId`, _ProjectId);
            localStorage.setItem(`${ProjectNo}-ProjectName`, _ProjectName);
            localStorage.setItem(`${ProjectNo}-WorkType`, _WorkType);       
        }                 
	}
	catch (e) {	
		console.log(e);	
	}
}

function fetchProjectTeam(method, serviceUrl, isAsync) {
    return new Promise((resolve, reject) => {
        var xhr = new XMLHttpRequest();
        xhr.open(method, serviceUrl, isAsync);

        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        const response = xhr.responseText; // Use xhr.responseText to access the response text
                        const parser = new DOMParser();
                        const xmlDoc = parser.parseFromString(response, "text/xml");
                        resolve(xmlDoc); // Resolve the promise with the parsed XML document
                    } catch (err) {
                        reject(new Error(`Failed to parse XML: ${err.message}`));
                    }
                } else {
                    reject(new Error(`Failed to get valid response: ${xhr.statusText}`));
                }
            }
        };

        xhr.setRequestHeader('Content-Type', 'text/xml');
        xhr.send();
    });
}

const appBar = async function(){

    let masterNavBar = $('#SuiteNavPlaceHolder');
    let childMenu = `<div class="sp-appBar" id="sp-appBar" role="navigation" aria-label="App bar" tabindex="-1">
                        <ul class="sp-appBar-linkContainer">`;
                        
    appBarItems.map(item=>{
          let editors = item.editors;
          let readers = item.readers;
          let className = item.iconTitle;
          
          childMenu += `<li class="sp-appBar-linkLi" data-automation-id="sp-appBar-linkLi-globalnav">
                                <div class="sp-appBar-linkLiDiv">
                                  <a class="sp-appBar-link ${className}" role="button" href="${item.redirectUrl}" onclick="openInCustomWindow(event)" style="display: flex; flex-direction: column; align-items: center; text-decoration: none;">
                                    <svg class= ${className} xmlns="http://www.w3.org/2000/svg" x="0px" y="0px" width="22" height="22" viewBox="${item.viewBox}">
                                      ${item.svgPath}                                  
                                    </svg>
                                  <span style="font-size: 10px; color: black; margin-top: 3px;">${item.iconTitle}</span>                                                             
                                  </a>                              
                                  <div class="sp-appBar-tooltipDiv" style="position: absolute; left: 100%; top: 50%; transform: translateX(10px) translateY(-50%);">
                                    <span class="sp-appBar-tooltip" role="tooltip" id="sp-appbar-tooltiplabel-globalnav">${item.tooltip}</span>
                                  </div>
                                </div>
                            </li>`;
    })
    childMenu += "</ul></div>";
    masterNavBar.append(childMenu);
}

function getClassName(module){
    let className = '';
    switch(module){
      case 'Roles':
          className = !_isPD && !_isPM ? 'dimmed-svg' : '';
      break;
      
      case 'D365':
        className = !_isPD && !_isPM ? 'dimmed-svg' : '';
        break;
    }
    return className
}

function getAppBarItemsForProjects(projectNumbers, customAppBarItems, commonAppBarItems) {
    // If a single project number is passed, convert it to an array
    if (!Array.isArray(projectNumbers)) {
      projectNumbers = [projectNumbers];
    }
  
    // Collect all project-specific app bar items
    const projectSpecificItems = projectNumbers.flatMap(projectNumber => customAppBarItems[projectNumber] || []);
  
    // Return combined common and project-specific app bar items
    return [...projectSpecificItems, ...commonAppBarItems ];
}
//#endregion