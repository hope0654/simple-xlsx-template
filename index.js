
var filesaver = require('file-saver');
var JSZip = require("jszip");
var htmlparser2 = require("htmlparser2");
var domhandler = require("domhandler");
var domutils = require('domutils')

function fetch_buf(url) {
    return new Promise(function(resolve, reject) {
        var req = new XMLHttpRequest()
        req.open('GET', url, true)
        req.responseType = 'arraybuffer'
        req.onload = function() { resolve(req.response) }
        req.onerror = reject
        req.send()
    })
}

function get_zip_file_string(zip, path) {
    return zip.file(path).async("string")
}

function get_zip_file_dom(zip, path) {
    return zip.file(path).async("string").then(xml_to_dom)
}

function set_zip_file_string(zip, path, content) {
    return zip.file(path, content)
}

function set_zip_file_dom(zip, path, dom) {
    return zip.file(path, dom_to_xml(dom))
}

function xml_to_dom(xml) {
    return new Promise(function(resolve, reject) {
        var handler = new domhandler.DomHandler(function (error, dom) {
            error ? reject(error) : resolve(dom)
        });

        var parser = new htmlparser2.Parser(handler, {
            xmlMode: true,
        });
        parser.write(xml);
        parser.end();
    })
}

function dom_to_xml(dom) {
    return domutils.getOuterHTML(dom, { xmlMode: true })
}

function match_dom(element, selector_object) {
    var result =  element.type === 'tag' && element.tagName === selector_object.tagName && selector_object.attribs.every(function(attrib) {
        var actual = element.attribs[attrib.name]
        if (actual === undefined) return false
        var expect = attrib.value
        return expect ? actual === expect : true
    })

    return result
}

function select_dom(dom, selector_string) {
    dom = Array.isArray(dom) ? dom : [dom]
    var selectors = selector_string.split(/\s+/g).map(function(selector) {
        var results = selector.match(/([\w:]+)(?:\[(.*?)(?:=(.*?))?\])?/)
        return {
            tagName: results[1],
            attribs: results[2] ? [{
                name: results[2],
                value: results[3],
            }] : [],
        }
    })

    for (var i = 0; i < selectors.length; i++) {
        var selector_object = selectors[i]
        dom = domutils.findAll(function(element) {
            return match_dom(element, selector_object)
        }, dom)
    }

    return dom
}

function select_all(dom, selector) {
    return select_dom(dom, selector)
}

function select_one(dom, selector) {
    var found = select_dom(dom, selector)[0]
    if (!found) throw new Error('selector ' + selector + ' not found')
    return found
}

function set_dom_text(dom, text) {
    if (dom && dom.children[0].type === 'text') {
        dom.children[0].data = text
    }
}

function get_dom_text(dom) {
    if (dom && dom.children[0].type === 'text') {
        return dom.children[0].data
    }
}

function set_dom_attribute(dom, name, value) {
    if (dom && dom.type === 'tag') {
        dom.attribs[name] = value
    }
}

function shared_string_manager(zip) {
    var shared_string_path = 'xl/sharedStrings.xml'
    var shared_string_dom = null

    var get_shared_string = function() {
        return Promise.resolve().then(function() {
            if (shared_string_dom) return shared_string_dom

            return get_zip_file_dom(zip, shared_string_path)
                .then(function(dom) {
                    shared_string_dom = dom
                    return shared_string_dom
                })
        })
    }

    var set_shared_string = function() {
        if (shared_string_dom) {
            return set_zip_file_dom(zip, shared_string_path, shared_string_dom)
        }
    }

    return {
        get: get_shared_string,
        set: set_shared_string,
    }
}

function path_relation_resolve(root, relative) {
    var root_parts = root.split('/')
    var relative_parts = relative.split('/')
    var dotdot_count = 0
    while (relative_parts[dotdot_count] === '..') dotdot_count++
    return root_parts.slice(0, -dotdot_count-1).concat(relative_parts.slice(dotdot_count)).join('/')
}


function get_relation_path(path) {
    var parts = path.split('/')
    var filename = parts[parts.length - 1] + '.rels'
    return parts.slice(0, -1).concat(['_rels', filename]).join('/')
}

function get_relations (zip, path) {
    var relation_path = get_relation_path(path)
    return Promise.resolve().then(function() {
        return get_zip_file_dom(zip, relation_path)
            .then(function(sheet_relations_dom) {
                var relationships_dom = select_all(sheet_relations_dom, 'Relationship')

                return relations = relationships_dom.reduce(function(m, relation_dom) {
                    var id = relation_dom.attribs.Id
                    var target = path_relation_resolve(path, relation_dom.attribs.Target)
                    m[id] = target
                    return m
                }, {})
            })
    })
}

function patch_sheet_chart(sheet_id, dom, cells, share) {
    var cf_doms = select_all(dom, 'c:ser c:f')

    for (var i = 0; i < cf_doms.length; i++) {
        var cf_dom = cf_doms[i]
        var cf_text = get_dom_text(cf_dom)
        var parts = cf_text.match(/(\w+\d+)!\$(\w+)\$(\d+):\$(\w+)\$(\d+)/)
        var sheet_idx = parts[1].toLowerCase()
        if (sheet_id !== sheet_idx) continue


        var start_col = parts[2]
        var start_row = +parts[3]
        var end_col = parts[4]
        var end_row = +parts[5]

        var cv_doms = select_all(cf_dom.parent, 'c:pt c:v')

        for (var j = 0; j < cv_doms.length; j++) {
            if (start_col === end_col) {
                var cell_id = start_col + (start_row + j).toString()
                var value = cells[cell_id]
                if (value) {
                    set_dom_text(cv_doms[j], value)
                }
            }
        }
    }
}

function patch_sheet_drawing(sheet_id, drawing_path, cells, share) {

    var zip = share.zip

    return Promise.all([
        get_zip_file_dom(zip, drawing_path),
        get_relations(zip, drawing_path),
        ])
    .then(function(results) {
        var drawing_file_dom = results[0]
        var drawing_relations = results[1]

        var charts_dom = select_all(drawing_file_dom, 'c:chart')

        return Promise.all(charts_dom.map(function(chart_dom) {
            var chart_rid = chart_dom.attribs['r:id']
            var chart_path = drawing_relations[chart_rid]

            return get_zip_file_dom(zip, chart_path)
            .then(function(chart_file_dom) {
                patch_sheet_chart(sheet_id, chart_file_dom, cells, share)
                return set_zip_file_dom(zip, chart_path, chart_file_dom)
            })
        }))
    })

}

function patch_sheet(sheet_id, dom, cells, share) {
    var cell_ids = Object.keys(cells)

    var populated_cells = {}

    return cell_ids.reduce(function(p, cell_id) {
        return p.then(function() {
            var cell_dom = select_one(dom, 'sheetData row c[r=' + cell_id + ']')
            var v_dom = select_one(cell_dom, 'v')
            var value = cells[cell_id]
            if (value === undefined || value === null) return

            if (cell_dom.attribs.t === 's') {
                return share.shared_string.get()
                    .then(function(share_string_dom) {

                        var v_text = get_dom_text(v_dom)
                        if (!v_text) return

                        var string_index = +v_text
                        var sit_dom = select_all(share_string_dom, 'sst si t')[string_index]

                        populated_cells[cell_id] = value
                        return set_dom_text(sit_dom, value)
                    })

            } else {
                populated_cells[cell_id] = value
                return set_dom_text(v_dom, value)
            }
        })
    }, Promise.resolve())
        .then(function() {
            return populated_cells
        })
}

function patch_and_save_sheet(sheet_id, cells, share) {
    var zip = share.zip
    var sheet_path = 'xl/worksheets/' + sheet_id + '.xml'

    return get_zip_file_dom(zip, sheet_path)
        .then(function(dom) {
            return patch_sheet(sheet_id, dom, cells, share)
                .then(function(populated_cells) {
                    if (Object.keys(populated_cells).length === 0) return
                    var drawing_dom = select_all(dom, 'drawing[r:id]')[0]
                    if (!drawing_dom) return

                    var drawing_rid = drawing_dom.attribs['r:id']

                    return get_relations(zip, sheet_path)
                        .then(function(relations) {

                            var drawing_path = relations[drawing_rid]
                            if (!drawing_path) return

                            return patch_sheet_drawing(sheet_id, drawing_path, populated_cells, share)
                        })

                })
                .then(function() {
                    return set_zip_file_dom(zip, sheet_path, dom)
                })
        })
}

function patch_xlsx(zip, data) {

    var share = {
        zip,
        shared_string: shared_string_manager(zip),
    }

    var sheet_ids = Object.keys(data)

    return sheet_ids.reduce(function(p, sheet_id) {
        return p.then(function() {
            var cells = data[sheet_id]
            if (Object.keys(cells).length === 0) return

            return patch_and_save_sheet(sheet_id, cells, share)
        })
    }, Promise.resolve())
        .then(share.shared_string.set)
}

function patch(url, data, name) {

    return fetch_buf(url).then(JSZip.loadAsync)
        .then(function(zip) {
            return Promise.resolve()
                .then(function() {
                    return patch_xlsx(zip, data)
                })
                .then(function() {
                    return zip.generateAsync({
                        type:"blob",
                        compression: "DEFLATE",
                    })
                })
                .then(function(blob) {
                    filesaver.saveAs(blob, name)
                })
        })
}

module.exports = patch

