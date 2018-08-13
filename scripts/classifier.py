import tensorflow as tf
import sys
import warnings
import os


os.environ['TF_CPP_MIN_LOG_LEVEL'] = 3

warnings.simplefilter("ignore")

image_path = sys.argv[1]
graph_path = "tf_files/retrained_graph.pb"
labels_path = "tf_files/retrained_labels.txt"

# Read input the image_data
image_data = tf.gfile.FastGFile(image_path, 'rb').read()

# Loads label file, strips off carriage return
label_lines = [line.rstrip() for line
               in tf.gfile.GFile(labels_path)]

# Load the Graph file
with tf.gfile.FastGFile(graph_path, 'rb') as f:
    graph_def = tf.GraphDef()
    graph_def.ParseFromString(f.read())
    _ = tf.import_graph_def(graph_def, name='')

# Feed the image_data as input to the graph and get first prediction
with tf.Session() as sess:
    softmax_tensor = sess.graph.get_tensor_by_name('final_result:0')
    predictions = sess.run(softmax_tensor,
                           {'DecodeJpeg/contents:0': image_data})
    # Sort to show labels of first prediction in order of confidence
    top_k = predictions[0].argsort()[-len(predictions[0]):][::-1]

# assign confidence values to images
    confidence = "Low"
    if predictions[0][top_k[0]] >= 0.85:
        confidence = "High"
    elif predictions[0][top_k[0]] > 0.75:
        confidence = "Medium"

    print('%s Confidence = %s Score = %.5f' % (label_lines[top_k[0]], confidence, predictions[0][top_k[0]]))

    # for i in top_k:
    #     classifier_label = label_lines[i]
    #     score = predictions[0][i]
    #     print('%s (Confidence = %.5f)' % (classifier_label, score))
    # print(label_lines[top_k[0]])
    # print(predictions[0][top_k[0]])
