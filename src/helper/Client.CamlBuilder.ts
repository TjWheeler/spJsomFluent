"use strict";   
/*  Script: Client.CamlBuilder
    By: Tim Wheeler
    Version: v7 - Mar 2019
        
    Notes: Date rage overlap query and scope not tested.
        
    TypeScript Example:
        var myAppData = myAppData || {};
(data => {
    var queryContext;
    var listName = "Tasks";
    data.query = () => {
        queryContext = {};
        var camlBuilder = new CamlBuilder();
        camlBuilder.addViewFields(["Title", "Priority"]);
        camlBuilder.addOrderBy("DueDate", false);
        camlBuilder.rowLimit = 4;
        camlBuilder.addDateTimeClause(CamlOperator.Geq, "DueDate", camlBuilder.addDaysToDate(new Date(), -7).toISOString(), false);
        camlBuilder.addDateTimeClause(CamlOperator.Leq, "DueDate", camlBuilder.addDaysToDate(new Date(), 7).toISOString(), false);
    
        var clientContext = SP.ClientContext.get_current();
        var ctx = clientContext;
        var list = ctx.get_web().get_lists().getByTitle(listName);
        var query = new SP.CamlQuery();
        query.set_viewXml(camlBuilder.viewXml);
        queryContext.items = list.getItems(query);
        ctx.load(queryContext.items);
        queryContext.list = list;
        queryContext.clientContext = clientContext;
        clientContext.executeQueryAsync(onQuerySucceeded, onQueryFailed);
    
    };
    
    function onQuerySucceeded() {
        var items = queryContext.items;
        var listItemInfo = "";
        var listItemEnumerator = items.getEnumerator();
        while (listItemEnumerator.moveNext()) {
            var listItem = listItemEnumerator.get_current();
    
            listItemInfo += '\nTitle: ' + listItem.get_item('Title');
            listItemInfo += ' Overdue: ' + listItem.get_item('Overdue');
        }
        console.log('success ' + listItemInfo);
    }
    function onQueryFailed(sender, args) {
        alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
    }
})(myAppData);
*/
export interface ICamlBuilder {
    query: string;
    viewXml: string;
    viewScope: string;
    viewFields: string;
    totalClauses: number;
    rowLimit: number;
    begin(requireAll: boolean);
    addNullClause(fieldRef: string);
    addTextClause(operation: CamlOperator, fieldRef: string, value: string);
    addNumberClause(operation: CamlOperator, fieldRef: string, value: number);
    addDateTimeClause(operation: CamlOperator, fieldRef: string, value: string, includeTime: boolean);
    addLookupClause(operation: CamlOperator, fieldRef: string, value: string, valueType: string);
    recurrenceQuery(overlapType: EventRecurranceOverlap);
    addViewField(fieldRef: string);
    addViewFields(fieldRefs: Array<string>);
    addGroupByField(fieldRef: string);
    addGroupByFields(fieldRefs: Array<string>);
    addOrderBy(fieldRef: string, ascending: boolean);
    addDaysToDate(date, days);
}

/// <summary>
///     CAML Operations
/// </summary>
export enum CamlOperator {
    Contains,
    BeginsWith,
    Eq,
    Geq,
    Leq,
    Lt,
    Gt,
    Neq,
    DateRangesOverlap,
    IsNotNull,
    IsNull
}

export enum EventRecurranceOverlap {
    Now,
    Today,
    Week,
    Month,
    Year
}
export enum AggregationType {
    SUM,
    COUNT,
    AVG,
    MAX,
    MIN,
    STDEV,
    VAR
}
export class StringBuilder {
    value: string = "";
    add(value: string) {
        this.value += value;
    }
    toString() {
        return this.value;
    }
}
export interface OdataItem {
    property: string;
    name: string;
    value: any;
}
export interface OdataPair {
    left: OdataPair;
    right: OdataPair;
    type: string;
    func: string;
    property: string;
    name: string;
    value: any;
    args: Array<OdataPair>;
}
export class CamlBuilder implements ICamlBuilder {
    constructor() {
        this.begin(true);
    }
    private camlClauses: Array<string> = [];
    private orderByFields: Array<OrderBy> = [];
    private requireAll: boolean = true;
    private viewFieldsString: string = "";
    private groupByFieldsString: string = "";
    private aggregationFieldsString: string = "";
    private recurrence: string;
    private directCaml: string;
    private queryOptions: string;
    rowLimit: number;
    viewScope: string;

    get query(): string {
        return this.buildQuery();
    }
    get viewFields(): string {
        return this.viewFieldsString;
    }
    get totalClauses(): number {
        return this.camlClauses.length;
    }
    setFilter(caml: string) {
            
        var hasWhereClause = caml.indexOf("<Where>") === 0;
        var sb = new StringBuilder();
        if (!hasWhereClause) {
            sb.add("<Where>");
        }
        sb.add(caml);
        if (!hasWhereClause) {
            sb.add("</Where>");
        }
        this.directCaml = sb.toString();
    }
    private buildQuery(): string {
        if (this.recurrence) {
            return `<Where>${this.recurrence}</Where>`;
        }
        var query = "";
        var openOperators = 0;
        if (this.directCaml) {
            query += this.directCaml;
        } else if (!this.directCaml && this.camlClauses.length > 0) {
            query += "<Where>";

            //When we have just one clause we can't use AND or OR.
            if (this.camlClauses.length > 1) {
                let totalCamlPairs = this.camlClauses.length - 1;
                while (totalCamlPairs > 0) {
                    query += (this.requireAll ? "<And>" : "<Or>");
                    totalCamlPairs--;
                    openOperators++;
                }
            }
            var clausesAdded = 0;

            for (let i = 0; i < this.camlClauses.length; i++) {
                const clause = this.camlClauses[i];
                query += clause;
                clausesAdded++;
                if (clausesAdded > 1) {
                    query += this.requireAll ? "</And>" : "</Or>";
                    openOperators--;
                }
            }
            query += "</Where>";
        }
        if (this.orderByFields.length > 0) {
            query += "<OrderBy>";
            for (let i = 0; i < this.orderByFields.length; i++) {
                var item = this.orderByFields[i];
                query += `<FieldRef Name="${item.fieldRef}" Ascending="${item.ascending ? "TRUE" : "FALSE"}" />`;
            }
            query += "</OrderBy>";
        }
        return query;
    }
    get viewXml(): string {
        var viewFields = "";
        if (this.viewFields) {
            viewFields = `<ViewFields>${this.viewFields}</ViewFields>`;
        }
        var groupBy = "";
        if (this.groupByFieldsString) {
            groupBy = `<GroupBy Collapse="TRUE" GroupLimit="1999">${this.groupByFieldsString}</GroupBy>`;
        }
        var orderBy = "";
        if (this.orderByFields && this.orderByFields.length) {
            orderBy = `<OrderBy>`;
            for (let i = 0; i < this.orderByFields.length; i++) {
                orderBy += `<FieldRef Name="${this.orderByFields[i].fieldRef}" Ascending="${this.orderByFields[i].ascending ? 'TRUE' : 'FALSE'}"/>`;
            }
            orderBy += `</OrderBy>`;
        }
        var aggregations = "";
        if (this.aggregationFieldsString) {
            aggregations = `<Aggregations Value="On">${this.aggregationFieldsString}</Aggregations>`;
        }
        var queryOptionText = "";
        if (this.queryOptions) {
            queryOptionText = `<QueryOptions>
                ${this.queryOptions}
            </QueryOptions>`;
        }
        return `<View${this.viewScope ? " " + this.viewScope : ""}>${viewFields}<Query>${groupBy}${(this.orderByFields && this.orderByFields.length > 0) || this.totalClauses > 0 || this.directCaml ? this.query : ""}${orderBy}</Query>${queryOptionText}${aggregations}${this.rowLimit ? "<RowLimit>" + this.rowLimit + "</RowLimit>" : ""}</View>`;
    }

    begin(requireAll: boolean) {
        this.camlClauses = [];
        this.requireAll = requireAll;
        this.recurrence = null;
        this.viewFieldsString = "";
        this.orderByFields = [];
        this.directCaml = null;
    }

    getNullClause(fieldRef: string) {
        var retVal = "";
        if (fieldRef) {
            retVal = `<IsNull><FieldRef Name="${fieldRef}" /></IsNull>`;
        }
        return retVal;
    }

    getClause(operation: CamlOperator, fieldRef: string, value: string, valueType: string): string {
        var retVal = `<${CamlOperator[operation]}><FieldRef Name="${fieldRef}" /><Value Type="${valueType}">${value}</Value></${CamlOperator[operation]}>`;
        return retVal;
    }
    getDateTimeClause(operation: CamlOperator, fieldRef: string, value: string, includeTime: boolean) {
        var retVal = `<${CamlOperator[operation]}><FieldRef Name="${fieldRef}" /><Value Type="DateTime" ${includeTime ? "IncludeTimeValue='TRUE'" : ""}>${value}</Value></${CamlOperator[operation]}>`;
        return retVal;
    }
    addNullClause(fieldRef: string) {
        this.camlClauses.push(this.getNullClause(fieldRef));
    }
    addTextClause(operation: CamlOperator, fieldRef: string, value: string) {
        this.camlClauses.push(this.getClause(operation, fieldRef, value, "Text"));
    }
    addBooleanClause(operation: CamlOperator, fieldRef: string, value: boolean) {
        this.camlClauses.push(this.getClause(operation, fieldRef, value ? "1" : "0", "Integer"));
    }
    addNumberClause(operation: CamlOperator, fieldRef: string, value: number) {
        this.camlClauses.push(this.getClause(operation, fieldRef, value.toString(), "Number")); //TODO: verify type
    }
    addDateTimeClause(operation: CamlOperator, fieldRef: string, value: string, includeTime: boolean) {
        this.camlClauses.push(this.getDateTimeClause(operation, fieldRef, value, includeTime));
    }

    addLookupClause(operation: CamlOperator, fieldRef: string, value: string, valueType: string) {
        var clause = `<${CamlOperator[operation]}><FieldRef Name="${fieldRef}" LookupId="True" /><Value Type="${valueType}">${value}</Value></${CamlOperator[operation]}>`;
        this.camlClauses.push(clause);
    }
    addAggregation(fieldRef: string, aggregationType: AggregationType) {
        this.aggregationFieldsString = this.aggregationFieldsString + `<FieldRef Name="${fieldRef.replace(" ", "_x0020_")}" Type="${AggregationType[aggregationType]}" />`;
    }
    recurrenceQuery(overlapType: EventRecurranceOverlap) {
        this.recurrence = `<DateRangesOverlap><FieldRef Name="EventDate"/><FieldRef Name="EndDate"/><FieldRef Name="RecurrenceID"/><Value><${EventRecurranceOverlap[overlapType]}/></Value></DateRangesOverlap>`;
        this.queryOptions += 
            `<CalendarDate>2018-01-01T12:00:00Z</CalendarDate>
            <RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion>
            <ExpandRecurrence>TRUE</ExpandRecurrence>`;
    }
    addViewField(fieldRef: string) {
        this.viewFieldsString = this.viewFieldsString + `<FieldRef Name="${fieldRef.replace(" ", "_x0020_")}"/>`;
    }
    addViewFields(fieldRefs: Array<string>) {
        for (var i = 0; i < fieldRefs.length; i++) {
            this.addViewField(fieldRefs[i]);
        }
    }
    addGroupByField(fieldRef: string) {
        this.groupByFieldsString = this.groupByFieldsString + `<FieldRef Name="${fieldRef.replace(" ", "_x0020_")}"/>`;
    }
    addGroupByFields(fieldRefs: Array<string>) {
        for (var i = 0; i < fieldRefs.length; i++) {
            this.addGroupByField(fieldRefs[i]);
        }
    }
    addOrderBy(fieldRef: string, ascending: boolean) {
        var orderBy = new OrderBy();
        orderBy.fieldRef = fieldRef;
        orderBy.ascending = ascending;
        this.orderByFields.push(orderBy);
    }
    addDaysToDate(date, days) {
        date.setDate(date.getDate() + days);
        return date;
    }
        
}
export class OrderBy {
    fieldRef: string;
    ascending: boolean;
}





