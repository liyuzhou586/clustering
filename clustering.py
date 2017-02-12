
# Hierarchial Clustering
from matplotlib import pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage
import numpy as np
import pandas as pd
import xlrd, json,os
from openpyxl import Workbook
from scipy.cluster.hierarchy import cophenet
from scipy.spatial.distance import pdist
from scipy.cluster.hierarchy import fcluster
from statistics import mean
from pylab import figure, axes, pie, title, show


def hc(path, sheet, rows):

    def write_file(data):                       # write the data to text for the register(s) to read
        f = open("center.txt", "a+")
        f.write("\n%s\n" % (data))
        f.close()

    def read_excel(path, sheet, numRows):       # load the excel file into a list
        book = xlrd.open_workbook(path)
        first_sheet = book.sheet_by_index(sheet)
        i = 1
        store = []
        while i < numRows:
            st = first_sheet.row_values(i)
            store.append(st)
            i += 1
        return store

    def fancy_dendrogram(*args, **kwargs):      # create dendrograms
        max_d = kwargs.pop('max_d', None)
        if max_d and 'color_threshold' not in kwargs:
            kwargs['color_threshold'] = max_d
        annotate_above = kwargs.pop('annotate_above', 0)

        ddata = dendrogram(*args, **kwargs)

        if not kwargs.get('no_plot', False):
            plt.title('Hierarchical Clustering Dendrogram (truncated)')
            plt.xlabel('sample index or (cluster size)')
            plt.ylabel('distance')
            for i, d, c in zip(ddata['icoord'], ddata['dcoord'], ddata['color_list']):
                x = 0.5 * sum(i[1:3])
                y = d[1]
                if y > annotate_above:
                    plt.plot(x, y, 'o', c=c)
                    plt.annotate("%.3g" % y, (x, y), xytext=(0, -5),
                                 textcoords='offset points',
                                 va='top', ha='center')
            if max_d:
                plt.axhline(y=max_d, c='k')
        return ddata

    if __name__ == "__main__":
        X = read_excel(path, sheet, rows)           # generate the linkage matrix
        Z = linkage(X, 'ward')
        c, coph_dists = cophenet(Z, pdist(X))

        # calculate full dendrogram
        plt.figure(figsize=(25, 10))
        plt.title('Hierarchical Clustering Dendrogram_%s' % (path.split('_')[1]))
        plt.xlabel('index')
        plt.ylabel('distance')
        dendrogram(
            Z,
            leaf_rotation=90.,  # rotates the x axis labels
            leaf_font_size=8.,  # font size for the x axis labels
        )
        plt.savefig('full_Dendogram_%s' % (path.split('_')[1]))
        # plt.title('Hierarchical Clustering Dendrogram (truncated) %s' % path)
        # plt.xlabel('index or (cluster size)')
        # plt.ylabel('distance')
        # dendrogram(
        #     Z,
        #     truncate_mode='lastp',  # show only the last p merged clusters
        #     p=12,  # show only the last p merged clusters
        #     leaf_rotation=90.,
        #     leaf_font_size=12.,
        #     show_contracted=True,  # to get a distribution impression in truncated branches
        # )
        # plt.show()
        # plt.savefig('trunc_dendogram')
        plt.figure(figsize=(25, 10))
        fancy_dendrogram(
            Z,
            truncate_mode='lastp',
            p=12,
            leaf_rotation=90.,
            leaf_font_size=12.,
            show_contracted=True,
            annotate_above=10,
            max_d=20,
        )
        max_d = 20  # Dendrogram Heights cut @ max_d
        plt.title('Hierarchical Clustering cut @ %d distance Dendrogram_%s' % (max_d, path.split('_')[1]))
        plt.savefig('cut_dendogram_%s' % (path.split('_')[1]))
        max_d = 20
        clusters = fcluster(Z, max_d, criterion='distance')
        clusters
        m = min(clusters)
        mx = max(clusters)
        clust = {}
        datapos = {}
        for i in range(m, mx + 1):
            clust[i] = []

        # path = "f5data_cleaned.xlsx"
        # Y = read_excel(path, 0, 101)                              #we can use the normalized values.

        for i in range(m, mx + 1):
            clust[i] = filter(lambda x: x == i, clusters)           # collect all the clusters- sort clusters

        for j in range(m, mx + 1):
            pos = [i for i, x in enumerate(clusters) if x == j]     # collect data points per position for a cluster
            temp = []
            for i in range(len(pos)):
                temp.append(X[pos[i]])                              # load the all row data inside cluster j
            datapos[j] = temp

        clustcenter = {}                                            # initialize the dictionary for cluster centers not required

        # create & write to an excel sheet
        wb = Workbook()
        ws0 = wb.create_sheet(0)  # Create the active worksheet
        ws0 = wb.active  # Grab the active worksheet
        ws0.append(
            ['ith Cluster', 'numofcluster', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
             'P', 'Q'])  # name the columns

        # find the cluster centers for each cluster by find the average of the column of each cluster data list
        t = 1
        for key, value in datapos.iteritems():
            st = []
            for row in range(len(datapos)):
                k = 0
                s = 0
                while k < len(value):
                    s += value[k][row]
                    k = k + 1
                st.append(s)

            write_file(json.dumps({key: [len(clust[key]), st]}))    # write to text file
            irow = [key, len(clust[key])] + st
            ws0.append([i for i in irow])                            # write to excel

        ws0.append(['Cophenet value =', c])                          # add cophenet values

        # write to disk
        wb.save('centers_%s.xlsx' % (path.split('_')[1]))            # Save the worksheet data
        print "Data for %s has a Cophenet value of %f" % (path.split('_')[1], c)  # cophenet goodness of the clusters


def iterate_files(num, period):
    p = "f5data_%s%d_norm.xlsx" %(period, num)
    return p


if __name__ == "__main__":
    for i in range(5):                                               #Specify the number of files
        path  = iterate_files(i+1, "day")                            #Specify the duration of the file
        print "Working on day %d data now" %(i+1)
        try:
            hc(path,0,289)                                           #Specify the number of sheet (e.g 0), rows per sheet
        except:
            print "file probably do not exist or rows exceeded in data excel"
