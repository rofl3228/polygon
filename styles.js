module.exports = (wb) => {
    return {
        common: wb.createStyle({
            font: {
                size: 9,
                family: 'roman',
            }
        }),
        header: wb.createStyle({
            alignment: {
                horizontal: 'center',
                vertical: 'center',
            },
            font: {
                size: 9,
                bold: true,
                vertAlign: 'center',
                family: 'roman',
            },
            fill: {
                type: 'pattern',
                patternType: "solid",
                bgColor: '#c6e0b4',
                fgColor: '#c6e0b4',
            },
            border: {
                right: {
                    style: 'medium',
                    color: '#000000',
                }
            }
        }),
        header_top_left: wb.createStyle({
            alignment: {
                horizontal: 'center',
                vertical: 'center',
                wrapText: true,
            },
            font: {
                size: 9,
                bold: true,
                family: 'roman',
                vertAlign: 'center',
            },
            fill: {
                type: 'pattern',
                patternType: "solid",
                bgColor: '#c6e0b4',
                fgColor: '#c6e0b4',
            },
            border: {
                left: {
                    style: 'medium',
                    color: '#000000'
                },
                right: {
                    style: 'medium',
                    color: '#000000'
                },
                top: {
                    style: 'medium',
                    color: '#000000'
                },
                bottom: {
                    style: 'medium',
                    color: '#000000'
                }
            },
        }),
        header_top: wb.createStyle({
            alignment: {
                horizontal: 'center',
                vertical: 'center',
            },
            font: {
                size: 9,
                bold: true,
                family: 'roman',
                vertAlign: 'center',
            },
            border: {
                right: {
                    style: 'medium',
                    color: '#000000',
                },
                top: {
                    style: 'medium',
                    color: '#000000',
                }
            },
            fill: {
                type: 'pattern',
                patternType: "solid",
                bgColor: '#c6e0b4',
                fgColor: '#c6e0b4',
            }
        }),
        title_green: wb.createStyle({
            alignment: {
                horizontal: 'center',
                vertical: 'center',
                wrapText: true,
            },
            font: {
                size: 9,
                bold: true,
                family: 'roman',
                vertAlign: 'center',
            },
            fill: {
                type: 'pattern',
                patternType: "solid",
                bgColor: '#27b665',
                fgColor: '#27b665',
            },
        }),
        title_red: wb.createStyle({
            alignment: {
                horizontal: 'center',
                vertical: 'center',
                wrapText: true,
            },
            font: {
                size: 9,
                bold: true,
                family: 'roman',
                vertAlign: 'center',
            },
            fill: {
                type: 'pattern',
                patternType: "solid",
                bgColor: '#ff6666',
                fgColor: '#ff6666',
            },
        }),
        border_left: wb.createStyle({
            border: {
                left: {
                    style: 'medium',
                    color: '#000000',
                }
            }
        }),
        border_bottom: wb.createStyle({
            font: {
                size: 9,
                bold: true,
                vertAlign: 'center',
            },
            border: {

                bottom: {
                    style: 'medium',
                    color: '#000000',
                }
            }
        }),
        right_end: wb.createStyle({
            border: {
                left: {
                    style: 'medium',
                    color: '#000000',
                }
            }
        })
    }
};
